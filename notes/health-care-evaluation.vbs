'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HEALTH CARE EVALUATION.vbs"
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
call changelog_update("05/31/2023", "Updated NOTES - Health Care Evaluation to include and reflect ex parte review process.", "Mark Riegel, Hennepin County")
call changelog_update("05/30/2023", "Updated NOTES - Health Care Evaluation to support recertification processing.##~####~##Added the MN Health Care Programs Renewal form as an option to select.##~##'Recertification' can be selected for each person with HC being processed.##~##", "Casey Love, Hennepin County")
call changelog_update("04/28/2023", "Updates to the script funcationality to support:##~## ##~## - PBEN infomration for the requirment of other programs.##~## - Indicate for a requirement to apply for Medicare.##~## - Selection of Major Program wlong with Basis of Eligibility.##~## - Additional fields for LTC specific information.##~## - Place to provide details of the AVS steps taken.##~## - If only one person on the case, the script will no longer require you select the household members.##~##", "Casey Love, Hennepin County")
call changelog_update("03/23/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'We need to load the information to read STAT from a class that is defined in its own script file
class_script_URL = script_repository & "misc/class-stat-detail.vbs"
If script_repository = "" Then
	run_locally = True
	class_script_URL = "C:\MAXIS-scripts\misc\class-stat-detail.vbs"
End If
IF on_the_desert_island = TRUE Then
	class_script_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\class-stat-detail.vbs"
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(class_script_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
Else
	Call run_from_GitHub(class_script_URL)
End If


'FUNCTIONS BLOCK ===========================================================================================================

function access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)
'reading information from AREP panel
	Call navigate_to_MAXIS_screen("STAT", "AREP")			'go to STAT/AREP

	EMReadScreen arep_name, 37, 4, 32						'reading the name to see if arep information exists
	arep_name = replace(arep_name, "_", "")
	If arep_name <> "" Then
		EMReadScreen arep_street_one, 22, 5, 32				'capturing information from the panel
		EMReadScreen arep_street_two, 22, 6, 32
		EMReadScreen arep_addr_city, 15, 7, 32
		EMReadScreen arep_addr_state, 2, 7, 55
		EMReadScreen arep_addr_zip, 5, 7, 64

		arep_street_one = replace(arep_street_one, "_", "")		'formatting information from the panel
		arep_street_two = replace(arep_street_two, "_", "")
		arep_addr_street = arep_street_one & " " & arep_street_two
		arep_addr_street = trim( arep_addr_street)
		arep_addr_city = replace(arep_addr_city, "_", "")
		arep_addr_state = replace(arep_addr_state, "_", "")
		arep_addr_zip = replace(arep_addr_zip, "_", "")

		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If arep_addr_state = left(state_item, 2) Then
				arep_addr_state = state_item
			End If
		Next

		EMReadScreen arep_phone_one, 14, 8, 34
		EMReadScreen arep_ext_one, 3, 8, 55
		EMReadScreen arep_phone_two, 14, 9, 34
		EMReadScreen arep_ext_two, 3, 8, 55

		arep_phone_one = replace(arep_phone_one, ")", "")
		arep_phone_one = replace(arep_phone_one, "  ", "-")
		arep_phone_one = replace(arep_phone_one, " ", "-")
		If arep_phone_one = "___-___-____" Then arep_phone_one = ""

		arep_phone_two = replace(arep_phone_two, ")", "")
		arep_phone_two = replace(arep_phone_two, "  ", "-")
		arep_phone_two = replace(arep_phone_two, " ", "-")
		If arep_phone_two = "___-___-____" Then arep_phone_two = ""

		arep_ext_one = replace(arep_ext_one, "_", "")
		arep_ext_two = replace(arep_ext_two, "_", "")

		EMReadScreen forms_to_arep, 1, 10, 45
		EMReadScreen mmis_mail_to_arep, 1, 10, 77
	End If
end function

function access_SWKR_panel(access_type, swkr_name, swkr_addr_street, swkr_addr_city, swkr_addr_state, swkr_addr_zip, swkr_phone_one, swkr_ext_one, notices_to_swkr_yn)
'reading information from the social worker (SWKR) panel
	Call navigate_to_MAXIS_screen("STAT", "SWKR")		'navigate to STAT/SWKR
	EMReadScreen swkr_name, 35, 6, 32
	swkr_name = replace(swkr_name, "_", "")
	If swkr_name <> "" Then								'if SWKR information exists, we read additional details
		EMReadScreen swkr_street_one, 22, 8, 32			'reading the information from SWKR
		EMReadScreen swkr_street_two, 22, 9, 32
		EMReadScreen swkr_addr_city, 15, 10, 32
		EMReadScreen swkr_addr_state, 2, 10, 54
		EMReadScreen swkr_addr_zip, 5, 10, 63

		swkr_street_one = trim(replace(swkr_street_one, "_", ""))		'format information read from SWKR
		swkr_street_two = trim(replace(swkr_street_two, "_", ""))
		swkr_addr_street = swkr_street_one & " " & swkr_street_two
		swkr_addr_street = trim( swkr_addr_street)
		swkr_addr_city = trim(replace(swkr_addr_city, "_", ""))
		swkr_addr_state = trim(replace(swkr_addr_state, "_", ""))
		swkr_addr_zip = trim(replace(swkr_addr_zip, "_", ""))

		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If swkr_addr_state = left(state_item, 2) Then
				swkr_addr_state = state_item
			End If
		Next

		EMReadScreen swkr_phone_one, 14, 12, 34
		EMReadScreen swkr_ext_one, 4, 12, 54

		swkr_phone_one = replace(swkr_phone_one, ")", "")
		swkr_phone_one = replace(swkr_phone_one, "  ", "-")
		swkr_phone_one = replace(swkr_phone_one, " ", "-")
		If swkr_phone_one = "___-___-____" Then swkr_phone_one = ""
		swkr_ext_one = trim(replace(swkr_ext_one, "_", ""))

		EMReadScreen notices_to_swkr_yn, 1, 15, 63
		notices_to_swkr_yn = trim(replace(notices_to_swkr_yn, "_", ""))
	End If
end function

function check_for_errors(eval_questions_clear)
'This is a function specific to this script to see if there are dialog errors that prevent us from moving forward in the script.
	For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)				'maandatory fields related to specific persons on the case from the first dialog
		If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
			If HEALTH_CARE_MEMBERS(HC_eval_process_const, the_memb) = "Select One..." Then err_msg = err_msg & "~!~" & "1 ^* Health Care Eval is at##~##   - Detail what type of evaluation is being cmopleted for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & ".##~##"
			If HEALTH_CARE_MEMBERS(HC_major_prog_const, selected_memb) <> "None" Then
				If HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, selected_memb) = "Select One..." Then err_msg = err_msg & "~!~" & "1 ^* MA Basis of Eligibility##~##   - Select what the Basis of Eligiblity of MA is for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & ".##~##"
			End If
			If HEALTH_CARE_MEMBERS(MSP_major_prog_const, selected_memb) <> "None" Then
				If HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, selected_memb) = "Select One..." Then err_msg = err_msg & "~!~" & "1 ^* MSP Basis of Eligibility##~##   - Select what the Basis of Eligiblity of MSP is for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & ".##~##"
			End If
			If HEALTH_CARE_MEMBERS(HC_major_prog_const, selected_memb) = "None" and HEALTH_CARE_MEMBERS(MSP_major_prog_const, selected_memb) = "None" Then err_msg = err_msg & "~!~" & "1 ^* HC/MSP Basis of Eligibility##~##   - At least one Major Program needs to be selected for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & ", indicate if MA/EMA/or a MSP is being assessed.##~##"
		End If
	Next
	If HC_form_name = "Breast and Cervical Cancer Coverage Group (DHS-3525)" Then		'handling for mandatory fields ONLY if MA-BC is being processed
		If trim(ma_bc_authorization_form) = "Select One..." Then err_msg = err_msg & "~!~" & "1 ^* Select authorization form needed##~##   - Select the form name needed for MA-BC Eligibility.##~##"
		If ma_bc_authorization_form_missing_checkbox = checked and IsDate(ma_bc_authorization_form_date) = True Then err_msg = err_msg & "~!~" & "1 ^* Check here if the form is NOT received and still required.##~##   - You checked the box indicating that the MA-BC authorization form was missing but entered a date for when the form was received."
		If ma_bc_authorization_form_missing_checkbox = unchecked and IsDate(ma_bc_authorization_form_date) = False Then err_msg = err_msg & "~!~" & "1 ^* Date Received (for MA-BC Authoriazation Form)##~##   - Enter the date the form for MA-BC Authorization was received."
	End If
	dlg_last_page_2_digits = left(last_page_numb&" ", 2)		'The dialog page needs to always be 2 digits or the functionality to display the errors has weird formatting

	'last page errors
	If app_sig_status = "Select One..." Then err_msg = err_msg & "~!~" & dlg_last_page_2_digits & "^* Has the Application been correctly Signed and Dated?##~##   - Select if all required signatures are on the application and correctly dated." & ".##~##"
	If app_sig_status = "No - Some applications or dates are missing" and trim(app_sig_notes) = "" THen err_msg = err_msg & "~!~" & dlg_last_page_2_digits & "^* If not, describe what is missing:##~##   - Since the application is not correctly signed/dated, enter the details of what is missing or incorrect." & ".##~##"

	For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
		If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
			If HEALTH_CARE_MEMBERS(hc_eval_status, the_memb) = "Select One..." Then err_msg = err_msg & "~!~" & dlg_last_page_2_digits & "^* Health Care Eval##~##   - Indicate the status of the Health Care Evaluation for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & ".##~##"
			If HEALTH_CARE_MEMBERS(hc_eval_status, the_memb) = "Incomplete - need additional verificaitons" and verifs_needed = "" Then err_msg = err_msg & "~!~" & dlg_last_page_2_digits & "^* Health Care Eval##~##   - You have indicated that the Health Care Evaluation for  MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & " is incomplete pending verifications but no verifications have been indicated. ##~##- Either update the status or press 'Update Verification' to document the details of the verifications needed.##~##"
			If HEALTH_CARE_MEMBERS(hc_eval_status, the_memb) = "Incomplete - other" and trim(HEALTH_CARE_MEMBERS(hc_eval_notes, the_memb)) = "" Then err_msg = err_msg & "~!~" & dlg_last_page_2_digits & "^* Evaluation Notes##~##   - Explain the details of the Health Care Evaluation Status for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & " as you have selected 'Other'.##~##- Add notes to 'Evaluation Notes' for further explanation.##~##"
		End if
	Next
end function

function display_errors(the_err_msg, execute_nav, show_err_msg_during_movement)
'function specific to this script that can display the errors in the err string with headers that identify the dialog page.
    If the_err_msg <> "" Then       'If the error message is blank - there is nothing to show.
        If left(the_err_msg, 3) = "~!~" Then the_err_msg = right(the_err_msg, len(the_err_msg) - 3)     'Trimming the message so we don't have a blank array item
        err_array = split(the_err_msg, "~!~")           'making the list of errors an array.

        error_message = ""                              'blanking out variables
        msg_header = ""
        for each message in err_array                   'going through each error message to order them and add headers'
			If ButtonPressed = completed_hc_eval_btn Then
	            current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing = trim(current_listing)
	            If current_listing <> msg_header Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
	                If current_listing = "1"  					Then tagline = ": HC MEMBs"        'Adding a specific tagline to the header for the errors
	                If current_listing = last_page_numb & ""  	Then tagline = ": App Info"
	                error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
	            End If
	            if msg_header = "" Then back_to_dialog = current_listing
	            msg_header = current_listing        'setting for the next loop

	            message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

	            error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
			Else
				If page_display = show_member_page Then page_to_review = "1"
				If page_display = last_page 	Then page_to_review = last_page_numb & ""

				current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing =  trim(current_listing)
				If current_listing = page_to_review Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
					If current_listing = "1"  					Then tagline = ": HC MEMBs"        'Adding a specific tagline to the header for the errors
					If current_listing = last_page_numb & "" 	Then tagline = ": App Info"
					If error_message = "" Then error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
					message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

					error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
				End If
			End If
        Next
		If error_message = "" then the_err_msg = ""

		show_msg = False
        If show_err_msg_during_movement = True Then show_msg = True
		If page_display = last_page AND (ButtonPressed <> completed_hc_eval_btn AND ButtonPressed <> next_btn AND ButtonPressed <> -1) Then show_msg = False

		If ButtonPressed = verif_button Then show_msg = False
		If ButtonPressed = clear_verifs_btn Then show_msg = False
		' If ButtonPressed = open_hsr_manual_transfer_page_btn Then show_msg = False
		If ButtonPressed >= 4000 Then show_msg = False
		For i = 0 to Ubound(HEALTH_CARE_MEMBERS, 2)
			If ButtonPressed = HEALTH_CARE_MEMBERS(pers_btn_one_const, i) Then show_msg = False
		Next

		If error_message = "" Then show_msg = False
		If ButtonPressed = completed_hc_eval_btn Then show_msg = True
		If page_display = show_pg_last Then
			If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
		End If

		If show_msg = True Then view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")
		If show_msg = False then the_err_msg = ""
        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE AND show_err_msg_during_movement = False Then
            If back_to_dialog = "1"  				Then ButtonPressed = hc_memb_btn         'This calls another function to go to the first dialog that had an error
            If back_to_dialog = last_page_numb & "" Then ButtonPressed = last_btn

            Call dialog_movement          'this is where the navigation happens
        End If
    End If
End Function

function define_main_dialog()
'this function is specific to this script to create the image of the dialog.
'This uses variables that are set to numbers to be equal to 'page_display'
'Each section of the if statements is the details of a specific dialog page.
'The container and buttons are defined for all the options to be the same (reducing the duplicate code)
	BeginDialog Dialog1, 0, 0, 555, 385, "Information for Health Care Evaluation"
	  ButtonGroup ButtonPressed
	  	'here starts the page specific display details
	    If page_display = show_member_page Then																	'MEMBER page
			GroupBox 10, 10, 465, 30, "Residents Requesting Health Care Coverage"
			x_pos = 10
			For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
				If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
                    If the_memb = selected_memb Then
    					Text x_pos+5, 25, 45, 10, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb)
    				Else
    					' PushButton 10, y_pos, 45, 10, "Person " & (the_memb + 1), HH_MEMB_ARRAY(button_one, the_memb)
						PushButton x_pos, 23, 40, 12, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb), HEALTH_CARE_MEMBERS(pers_btn_one_const, the_memb)
    				End If
    				x_pos = x_pos + 45
                End If
			Next
			' PushButton 50, 25, 40, 15, "MEMB 01", Button5
			GroupBox 10, 45, 465, 310, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, selected_memb) & " - PMI: " & HEALTH_CARE_MEMBERS(pmi_const, selected_memb)
			Text 250, 45, 200, 10, "Current MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " Health Care Status: " & HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, selected_memb)

			Text 300, 60, 80, 10, "Health Care Eval is at "
			DropListBox 380, 55, 85, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"No Evaluation Needed", HEALTH_CARE_MEMBERS(HC_eval_process_const, selected_memb)
			Text 20, 75, 180, 10, "Member: " & HEALTH_CARE_MEMBERS(full_name_const, selected_memb)
			Text 35, 85, 75, 10, "AGE: " & HEALTH_CARE_MEMBERS(age_const, selected_memb)
			Text 215, 75, 75, 10, "SSN: " & HEALTH_CARE_MEMBERS(ssn_const, selected_memb)
			Text 215, 85, 75, 10, "DOB: " & HEALTH_CARE_MEMBERS(dob_const, selected_memb)
			Text 320, 75, 110, 10, " Application Date: " & HEALTH_CARE_MEMBERS(hc_appl_date_const, selected_memb)
			Text 315, 85, 95, 10, "Coverage Request: " & HEALTH_CARE_MEMBERS(hc_cov_date_const, selected_memb)


			' Text 20, 295, 400, 10, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, selected_memb) & " - PMI: " & HEALTH_CARE_MEMBERS(pmi_const, selected_memb)

			' GroupBox 10, 220, 465, 60, "Medical Assistance"
			DropListBox 20, 100, 35, 45, "MA"+chr(9)+"EMA"+chr(9)+"None", HEALTH_CARE_MEMBERS(HC_major_prog_const, selected_memb)
			Text 55, 105, 100, 10, " - HC Program - Basis of ELIG:"
			' Text 65, 110, 60, 10, "Basis of ELIG:"
			DropListBox 160, 100, 90, 45, "Select One..."+chr(9)+ma_basis_of_elig_list, HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, selected_memb)
			Text 255, 105, 65, 10, "Prog/Basis Notes:"
			EditBox 315, 100, 150, 15, HEALTH_CARE_MEMBERS(MA_basis_notes_const, selected_memb)


			' GroupBox 10, 285, 465, 60, "Medicare Savings Programs"
			DropListBox 20, 120, 35, 45, "None"+chr(9)+"QMB"+chr(9)+"SLMB"+chr(9)+"QI1", HEALTH_CARE_MEMBERS(MSP_major_prog_const, selected_memb)
			Text 55, 125, 100, 10, " - MSP Program - Basis of ELIG:"
			' Text 20, 130, 70, 10, "MSP Basis of ELIG:"
			DropListBox 160, 120, 90, 45, "Select One..."+chr(9)+msp_basis_of_elig_list, HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, selected_memb)
			Text 255, 125, 65, 10, "Prog/Basis Notes:"
			EditBox 315, 120, 150, 15, HEALTH_CARE_MEMBERS(MSP_basis_notes_const, selected_memb)

			y_pos = 140
			If HC_form_name = "Breast and Cervical Cancer Coverage Group (DHS-3525)" Then
				GroupBox 10, y_pos, 465, 50,"MA for Breast/Cervical Cancer Form Required"
				y_pos = y_pos + 15
				Text 20, y_pos+5, 115, 10, "Select authorization form needed:"
				DropListBox 135, y_pos, 150, 45, "Select One..."+chr(9)+"SAGE Enrollment Form"+chr(9)+"Screen Our Circle Form"+chr(9)+"Certification of Further Treatment Required", ma_bc_authorization_form
				Text 290, y_pos+5, 55, 10, "date received"
				EditBox 345, y_pos, 50, 15, ma_bc_authorization_form_date
				CheckBox 135, y_pos+20, 200, 10,"Check here if the form is NOT received and still required.", ma_bc_authorization_form_missing_checkbox
				y_pos = y_pos + 40
			End If

			If HEALTH_CARE_MEMBERS(DISA_exists_const, selected_memb) = True Then
				Text 20, y_pos, 200, 10, "DISA    -    Start date: " & HEALTH_CARE_MEMBERS(DISA_start_date_const, selected_memb) & "   -   End Date: " & HEALTH_CARE_MEMBERS(DISA_end_date_const, selected_memb)
				Text 250, y_pos, 200, 10, "HC DISA Status: " & HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, selected_memb)
				y_pos = y_pos + 10
				If HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) <> "" Then Text 55, y_pos, 230, 10, "Certification   -   Start date: " & HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) & " - End Date: " & HEALTH_CARE_MEMBERS(DISA_cert_end_const, selected_memb)
				If HEALTH_CARE_MEMBERS(DISA_cert_start_const, selected_memb) = "" Then Text 55, y_pos, 230, 10, "Certification   -   NO CERTIFICATION DATES Entered"
				Text 285, y_pos, 150, 10, "  Verif: " & HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, selected_memb)
				y_pos = y_pos + 10
				If HEALTH_CARE_MEMBERS(DISA_waiver_info_const, selected_memb) <> "" Then
					Text 55, y_pos, 200, 10, "LTC Waiver: " & HEALTH_CARE_MEMBERS(DISA_waiver_info_const, selected_memb)
				Else
					Text 55, y_pos, 350, 10, "NO Waiver indicated. IF a WAIVER is being requested, add details in the NOTES section here."
				End If
				y_pos = y_pos + 15
				Text 55, y_pos, 45, 10, "DISA Notes:"
				EditBox 100, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(DISA_notes_const, selected_memb)
				y_pos = y_pos + 10
			Else
				Text 20, y_pos, 355, 10, "DISA   -   No DISA Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
				y_pos = y_pos + 10
			End If
			y_pos = y_pos + 5

			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) Then
					If STAT_INFORMATION(month_ind).stat_emma_exists(each_memb) = True Then
						Text 20, y_pos, 200, 10, "EMMA  -  Medical Emergency: " & STAT_INFORMATION(month_ind).stat_emma_med_emer_info(each_memb)
						Text 250, y_pos, 200, 10, "Health Consequence: " & STAT_INFORMATION(month_ind).stat_emma_health_cons_info(each_memb)
						y_pos = y_pos + 10
						Text 55, y_pos, 200, 10, "Begin Date: " & STAT_INFORMATION(month_ind).stat_emma_begin_date(each_memb) & " - End Date: " & STAT_INFORMATION(month_ind).stat_emma_end_date(each_memb)
						Text 250, y_pos, 200, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_emma_verif_info(each_memb)
						y_pos = y_pos + 15
						Text 55, y_pos, 45, 10, "EMMA Notes:"
						' EditBox 100, y_pos-5, 365, 15, STAT_INFORMATION(month_ind).stat_emma_notes(each_memb)
						EditBox 100, y_pos-5, 365, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_emma_notes(each_memb))
						y_pos = y_pos + 15
					End If
				End If
			Next

			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) Then
					If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = False Then
						Text 20, y_pos, 380, 10, "IMIG    -    No Immigration Information exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
					End if
					If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = True Then
						Text 20, y_pos, 380, 10, "IMIG    -    Immigration information for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " is on the IMIG page to the right."
					End if
					y_pos = y_pos + 15
				End If
			Next

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
				Text 20, y_pos, 160, 10, "MEDI   -    No MEDI Panel Exists for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
				CheckBox 180, y_pos, 190, 10, "Check here if an application for Medicare is required.", HEALTH_CARE_MEMBERS(MEDI_application_requred_checkbox_const, selected_memb)
				' y_pos = y_pos + 10
				Text 370, y_pos, 50, 10, "Referral Date:"
				EditBox 420, y_pos-5, 50, 15, HEALTH_CARE_MEMBERS(MEDI_referral_date_const, selected_memb)
				' y_pos = y_pos + 15
				y_pos = y_pos + 10
				'TODO - add in no MEDI
			End If
			y_pos = y_pos + 5

			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) Then
					If STAT_INFORMATION(month_ind).stat_pben_exists(each_memb) = True Then
						Text 20, y_pos, 380, 10, "PBEN   -   Potential Benefits for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb) & " listed:"
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_pben_type_code_one(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_one(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_one(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_one(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_one(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_one(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_one(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_one(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_one(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_one(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						If STAT_INFORMATION(month_ind).stat_pben_type_code_two(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_two(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_two(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_two(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_two(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_two(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_two(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_two(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_two(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_two(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						If STAT_INFORMATION(month_ind).stat_pben_type_code_three(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_three(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_three(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_three(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_three(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_three(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_three(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_three(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_three(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_three(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						If STAT_INFORMATION(month_ind).stat_pben_type_code_four(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_four(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_four(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_four(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_four(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_four(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_four(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_four(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_four(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_four(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						If STAT_INFORMATION(month_ind).stat_pben_type_code_five(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_five(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_five(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_five(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_five(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_five(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_five(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_five(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_five(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_five(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						If STAT_INFORMATION(month_ind).stat_pben_type_code_six(each_memb) <> "" Then
							Text 55, y_pos, 410, 10,  STAT_INFORMATION(month_ind).stat_pben_type_info_six(each_memb) & "   -   Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_six(each_memb) & "   -   Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_six(each_memb)
							y_pos = y_pos + 10
							date_detail = ""
							If STAT_INFORMATION(month_ind).stat_pben_referral_date_six(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_six(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_date_applied_six(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_six(each_memb) & "   -   "
							If STAT_INFORMATION(month_ind).stat_pben_iaa_date_six(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_six(each_memb) & "   -   "
							date_detail = left(date_detail, len(date_detail)-7)
							Text 80, y_pos, 350, 10,  date_detail
							y_pos = y_pos + 10
						End If
						y_pos = y_pos + 5
						Text 55, y_pos, 45, 10, "PBEN Notes:"
						EditBox 100, y_pos-5, 365, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_pben_notes(each_memb))
						y_pos = y_pos + 10
					Else
						Text 20, y_pos, 380, 10, "PBEN   -   No Potential Benefits indicated for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, selected_memb)
						y_pos = y_pos + 10
					End if
				End If
			Next





			Text 505, 17, 55, 13, "HC MEMBs"
		ElseIf page_display = show_jobs_page Then															'JOBS Page
			grp_len = 5
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
			Next
			If grp_len = 5 Then grp_len = 100

			GroupBox 10, 10, 465, grp_len, "JOBS Income"
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_start_date(each_memb)
					Text 235, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_one_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, y_pos+1, 430, 20, "Pay Detail"
					y_pos = y_pos + 9
					Text 100, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_one_main_pay_freq(each_memb)
					Text 235, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_one_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 16
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_one_notes(each_memb))
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_start_date(each_memb)
					Text 235, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_two_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, y_pos+1, 430, 20, "Pay Detail"
					y_pos = y_pos + 9
					Text 100, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_two_main_pay_freq(each_memb)
					Text 235, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_two_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 16
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_two_notes(each_memb))
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_start_date(each_memb)
					Text 235, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_three_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, y_pos+1, 430, 20, "Pay Detail"
					y_pos = y_pos + 9
					Text 100, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_three_main_pay_freq(each_memb)
					Text 235, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_three_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 16
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_three_notes(each_memb))
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_start_date(each_memb)
					Text 235, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_four_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, y_pos+1, 430, 20, "Pay Detail"
					y_pos = y_pos + 9
					Text 100, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_four_main_pay_freq(each_memb)
					Text 235, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_four_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 16
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_four_notes(each_memb))
					y_pos = y_pos + 25
				End If
				If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Employed at " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb)
					y_pos = y_pos + 15
					Text 30, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_start_date(each_memb)
					Text 235, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_end_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_jobs_five_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then Text 240, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					y_pos = y_pos + 10

					GroupBox 30, y_pos+1, 430, 20, "Pay Detail"
					y_pos = y_pos + 9
					Text 100, y_pos, 100, 10, "Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_five_main_pay_freq(each_memb)
					Text 235, y_pos, 175, 10, "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_five_health_care_income_pay_day(each_memb) & " per pay date"
					y_pos = y_pos + 16
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_five_notes(each_memb))
					y_pos = y_pos + 25
				End If
			Next

			If y_pos = 25 Then
				Text 20, 25, 350, 10, "NO JOBS panels have been entered in the csae file for the selected members."
				Text 50, 35, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				Text 20, 50, 350, 20, "If there is income from a job that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				Text 20, 70, 350, 10, "CASE/NOTE will indicate NO JOBS, add any notes here that are relevant:"
				EditBox 20, 80, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_general_notes)
			End If
			'TODO - add STWK information


			Text 500, 32, 55, 13, "JOBS Income"
		ElseIf page_display = show_busi_page Then															'BUSI Page
			grp_len = 15
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
				If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
					grp_len = grp_len + 90
				End If
			Next
			If grp_len = 15 Then grp_len = 100
			GroupBox 10, 10, 465, grp_len, "BUSI Income"

			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_one_type_info(each_memb)
					y_pos = y_pos + 10
					Text 280, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_busi_one_inc_start_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_one_inc_end_date(each_memb) <> "" Then Text 280, y_pos + 10, 115, 10, " Income End Date: " & STAT_INFORMATION(month_ind).stat_busi_one_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb) <> "" Then
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: B"
						If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_net_inc(each_memb) Then Text 30, y_pos, 160, 10, "HC Calculation Method: A and B"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					Else
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: A"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_one_hc_a_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End if
					y_pos = y_pos + 10
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_one_notes(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_two_type_info(each_memb)
					y_pos = y_pos + 10
					Text 280, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_busi_two_inc_start_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_two_inc_end_date(each_memb) <> "" Then Text 280, y_pos + 10, 115, 10, " Income End Date: " & STAT_INFORMATION(month_ind).stat_busi_two_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb) <> "" Then
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: B"
						If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_net_inc(each_memb) Then Text 30, y_pos, 160, 10, "HC Calculation Method: A and B"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					Else
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: A"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_two_hc_a_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End if
					y_pos = y_pos + 10
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_two_notes(each_memb))
				End If
				If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 235, y_pos, 230, 10, "Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_three_type_info(each_memb)
					y_pos = y_pos + 10
					Text 280, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_busi_three_inc_start_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_three_inc_end_date(each_memb) <> "" Then Text 280, y_pos + 10, 115, 10, " Income End Date: " & STAT_INFORMATION(month_ind).stat_busi_three_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb) <> "" Then
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: B"
						If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_net_inc(each_memb) Then Text 30, y_pos, 160, 10, "HC Calculation Method: A and B"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					Else
						Text 30, y_pos, 175, 10, "NET Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_net_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "Gross Income: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_gross_inc(each_memb)
						y_pos = y_pos + 10
						Text 105, y_pos, 105, 10, "  -   Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_expenses(each_memb)
						y_pos = y_pos + 10
						Text 30, y_pos, 160, 10, "HC Calculation Method: A"
						Text 235, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_income_verif_info(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_busi_three_hc_a_income_verif_code(each_memb) = "N" Then Text 275, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End if
					y_pos = y_pos + 10
					Text 30, y_pos+5, 50, 10, "Income Notes:"
					EditBox 80, y_pos, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_three_notes(each_memb))
				End If
			Next

			If y_pos = 25 Then
				Text 20, 25, 350, 10, "NO BUSI panels have been entered in the csae file for the selected members."
				Text 50, 35, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				Text 20, 50, 350, 20, "If there is income from self employment that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				Text 20, 70, 350, 10, "CASE/NOTE will indicate NO SELF EMPLOYMENT, add any notes here that are relevant:"
				EditBox 20, 80, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_general_notes)
			End If

			Text 500, 47, 55, 13, "BUSI Income"
		ElseIf page_display = show_unea_page Then															'UNEA Page
			grp_len = 15
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
					grp_len = grp_len + 45
					If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then grp_len = grp_len + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
					grp_len = grp_len + 45
					If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then grp_len = grp_len + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
					grp_len = grp_len + 45
					If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then grp_len = grp_len + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
					grp_len = grp_len + 45
					If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then grp_len = grp_len + 10
				End If
				If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
					grp_len = grp_len + 45
					If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then grp_len = grp_len + 10
				End If
			Next
			If grp_len = 15 Then grp_len = 100
			GroupBox 10, 10, 465, grp_len, "UNEA Income"
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_one_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_start_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_one_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then
						y_pos = y_pos + 10
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End If
					y_pos = y_pos + 15
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_one_notes(each_memb))
					' y_pos = y_pos + 10
					y_pos = y_pos + 20
				End If
				If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_two_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_start_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_two_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then
						y_pos = y_pos + 10
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End If
					y_pos = y_pos + 15
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_two_notes(each_memb))
					' y_pos = y_pos + 10
					y_pos = y_pos + 20
				End If
				If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_three_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_start_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_three_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then
						y_pos = y_pos + 10
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End If
					y_pos = y_pos + 15
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_three_notes(each_memb))
					' y_pos = y_pos + 10
					y_pos = y_pos + 20
				End If
				If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_four_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_start_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_four_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then
						y_pos = y_pos + 10
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End If
					y_pos = y_pos + 15
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_four_notes(each_memb))
					' y_pos = y_pos + 10
					y_pos = y_pos + 20
				End If
				If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
					grp_len = grp_len + 85
					Text 20, y_pos, 150, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 170, y_pos, 175, 10, "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_five_prosp_monthly_gross_income(each_memb)
					Text 320, y_pos, 115, 10, "Income Start Date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_start_date(each_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "Income type: " & STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb)
					Text 170, y_pos, 185, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_unea_five_verif_info(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb) <> "" Then Text 320, y_pos, 115, 10, "Income End Date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then
						y_pos = y_pos + 10
						Text 170, y_pos, 185, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
					End If
					y_pos = y_pos + 15
					Text 30, y_pos, 50, 10, "Income Notes:"
					EditBox 80, y_pos-5, 380, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_five_notes(each_memb))
					' y_pos = y_pos + 10
					y_pos = y_pos + 20
				End If
			Next

			If y_pos = 25 Then
				Text 20, 25, 350, 10, "NO UNEA panels have been entered in the csae file for the selected members."
				Text 50, 35, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				Text 20, 50, 350, 20, "If there is income from an unearned income source that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				Text 20, 70, 350, 10, "CASE/NOTE will indicate NO UNEARNED INCOME, add any notes here that are relevant:"
				EditBox 20, 80, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_general_notes)
			End If

			Text 500, 62, 55, 13, "UNEA Income"
		ElseIf page_display = show_asset_page Then															'Account Page

			' grp_len = 10
			' For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			' 	If STAT_INFORMATION(month_ind).stat_cash_asset_panel_exists(each_memb) = True Then
			' 		grp_len = grp_len + 45
			' 		If STAT_INFORMATION(month_ind).stat_cash_exists(each_memb) = True Then grp_len = grp_len + 15
			' 		If STAT_INFORMATION(month_ind).stat_acct_one_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_acct_two_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_acct_three_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_acct_four_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_acct_five_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_secu_one_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_secu_two_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_secu_three_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_secu_four_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_secu_five_exists(each_memb) = True Then
			' 			grp_len = grp_len + 25
			' 		End If
			' 	End If
			' Next
			' If grp_len = 10 Then grp_len = 100

			GroupBox 10, 10, 465, 65, "AVS Forms and Actions"
			Text 20, 25, 110, 10, "Status of AVS Authorization Form: "
			DropListBox 130, 20, 200, 45, avs_form_status_list, avs_form_status
			Text 342, 25, 20, 10, "Notes:"
			EditBox 365, 20, 105, 15, avs_form_notes
			Text 130, 35, 350, 10, "* Selecting 'Incomplete' or 'No Form Received' will add AVS Form Information to the List of Verifications."
			Text 20, 45, 150, 10, "AVS Portal Submission Actions and Notes:"
			EditBox 20, 55, 450, 15, avs_portal_notes

			y_pos = 90
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_cash_asset_panel_exists(each_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_cash_exists(each_memb) = True Then
						Text 25, y_pos, 115, 10, "CASH   -   Amount: $ " & STAT_INFORMATION(month_ind).stat_cash_balance(each_memb)
						y_pos = y_pos + 10
					End If

					If STAT_INFORMATION(month_ind).stat_acct_one_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "ACCT   -   Location: " & STAT_INFORMATION(month_ind).stat_acct_one_location(each_memb)
						Text 200, y_pos, 260, 10, "Account Type: " & STAT_INFORMATION(month_ind).stat_acct_one_type_detail(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 75, 10, "Balance: $ " & STAT_INFORMATION(month_ind).stat_acct_one_balance(each_memb)
						Text 135, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_acct_one_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_acct_one_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_acct_one_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If
					If STAT_INFORMATION(month_ind).stat_acct_two_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "ACCT   -   Location: " & STAT_INFORMATION(month_ind).stat_acct_two_location(each_memb)
						Text 200, y_pos, 260, 10, "Account Type: " & STAT_INFORMATION(month_ind).stat_acct_two_type_detail(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 75, 10, "Balance: $ " & STAT_INFORMATION(month_ind).stat_acct_two_balance(each_memb)
						Text 135, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_acct_two_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_acct_two_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_acct_two_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_acct_three_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "ACCT   -   Location: " & STAT_INFORMATION(month_ind).stat_acct_three_location(each_memb)
						Text 200, y_pos, 260, 10, "Account Type: " & STAT_INFORMATION(month_ind).stat_acct_three_type_detail(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 75, 10, "Balance: $ " & STAT_INFORMATION(month_ind).stat_acct_three_balance(each_memb)
						Text 135, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_acct_three_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_acct_three_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_acct_three_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_acct_four_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "ACCT   -   Location: " & STAT_INFORMATION(month_ind).stat_acct_four_location(each_memb)
						Text 200, y_pos, 260, 10, "Account Type: " & STAT_INFORMATION(month_ind).stat_acct_four_type_detail(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 75, 10, "Balance: $ " & STAT_INFORMATION(month_ind).stat_acct_four_balance(each_memb)
						Text 135, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_acct_four_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_acct_four_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_acct_four_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_acct_five_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "ACCT   -   Location: " & STAT_INFORMATION(month_ind).stat_acct_five_location(each_memb)
						Text 200, y_pos, 260, 10, "Account Type: " & STAT_INFORMATION(month_ind).stat_acct_five_type_detail(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 75, 10, "Balance: $ " & STAT_INFORMATION(month_ind).stat_acct_five_balance(each_memb)
						Text 135, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_acct_five_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_acct_five_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_acct_five_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
					End If

					If STAT_INFORMATION(month_ind).stat_secu_one_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "SECU   -   Name: " & STAT_INFORMATION(month_ind).stat_secu_one_name(each_memb)
						Text 200, y_pos, 160, 10, "Security Type: " & STAT_INFORMATION(month_ind).stat_secu_one_type_detail(each_memb)
						Text 360, y_pos, 110, 10, "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_one_cash_value(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 80, 10, "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_one_face_value(each_memb)
						Text 140, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_secu_one_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_secu_one_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_secu_one_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If
					If STAT_INFORMATION(month_ind).stat_secu_two_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "SECU   -   Name: " & STAT_INFORMATION(month_ind).stat_secu_two_name(each_memb)
						Text 200, y_pos, 160, 10, "Security Type: " & STAT_INFORMATION(month_ind).stat_secu_two_type_detail(each_memb)
						Text 360, y_pos, 110, 10, "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_two_cash_value(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 80, 10, "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_two_face_value(each_memb)
						Text 140, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_secu_two_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_secu_two_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_secu_two_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If
					If STAT_INFORMATION(month_ind).stat_secu_three_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "SECU   -   Name: " & STAT_INFORMATION(month_ind).stat_secu_three_name(each_memb)
						Text 200, y_pos, 160, 10, "Security Type: " & STAT_INFORMATION(month_ind).stat_secu_three_type_detail(each_memb)
						Text 360, y_pos, 110, 10, "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_three_cash_value(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 80, 10, "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_three_face_value(each_memb)
						Text 140, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_secu_three_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_secu_three_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_secu_three_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If
					If STAT_INFORMATION(month_ind).stat_secu_four_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "SECU   -   Name: " & STAT_INFORMATION(month_ind).stat_secu_four_name(each_memb)
						Text 200, y_pos, 160, 10, "Security Type: " & STAT_INFORMATION(month_ind).stat_secu_four_type_detail(each_memb)
						Text 360, y_pos, 110, 10, "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_four_cash_value(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 80, 10, "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_four_face_value(each_memb)
						Text 140, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_secu_four_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_secu_four_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_secu_four_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If
					If STAT_INFORMATION(month_ind).stat_secu_five_exists(each_memb) = True Then
						Text 25, y_pos, 175, 10, "SECU   -   Name: " & STAT_INFORMATION(month_ind).stat_secu_five_name(each_memb)
						Text 200, y_pos, 160, 10, "Security Type: " & STAT_INFORMATION(month_ind).stat_secu_five_type_detail(each_memb)
						Text 360, y_pos, 110, 10, "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_five_cash_value(each_memb)
						y_pos = y_pos + 10
						Text 58, y_pos, 80, 10, "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_five_face_value(each_memb)
						Text 140, y_pos, 60, 10, "as of " & STAT_INFORMATION(month_ind).stat_secu_five_as_of_date(each_memb)
						Text 205, y_pos, 115, 10, " Verification: " & STAT_INFORMATION(month_ind).stat_secu_five_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_secu_five_verif_code(each_memb) = "N" Then Text 325, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10

					End If

					Text 25, y_pos+5, 50, 10, "Asset Notes:"
					EditBox 75, y_pos, 395, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_asset_notes(each_memb))
					y_pos = y_pos + 25
				End If

				'TODO - DEAL WITH OTHR panel
			Next
			If y_pos = 90 Then
				Text 20, y_pos, 350, 10, "NO CASH/ACCT/SECU panels have been entered in the csae file for the selected members."
				y_pos = y_pos + 10
				Text 50, y_pos, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				y_pos = y_pos + 15
				Text 20, y_pos, 350, 20, "If there are liquid assets that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				y_pos = y_pos + 20
				Text 20, y_pos, 350, 10, "CASE/NOTE will indicate NO ACCOUNTS, add any notes here that are relevant:"
				y_pos = y_pos + 10
				EditBox 20, y_pos, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_acct_general_notes)
				y_pos = y_pos + 20
			End If

			grp_len = y_pos
			grp_len = grp_len - 80
			GroupBox 10, 80, 465, grp_len, "Assets"
			' GroupBox 10, 10, 465, grp_len, "Vehicles and Real Estate"

			Text 510, 77, 55, 13, "Assets"
		ElseIf page_display = show_cars_rest_page Then															'Cars ad Real Estate Page
			cars_exists = False
			rest_exists = False
			' For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			' 	If STAT_INFORMATION(month_ind).stat_cars_exists_for_member(each_memb) = True Then
			' 		grp_len = grp_len + 45
			' 		cars_exists = True
			' 		If STAT_INFORMATION(month_ind).stat_cars_one_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_cars_two_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_cars_three_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 	End If
			' 	If STAT_INFORMATION(month_ind).stat_rest_exists_for_member(each_memb) = True Then
			' 		rest_exists = True
			' 		grp_len = grp_len + 45
			' 		If STAT_INFORMATION(month_ind).stat_rest_one_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_rest_two_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 		If STAT_INFORMATION(month_ind).stat_rest_three_exists(each_memb) = True Then
			' 			grp_len = grp_len + 35
			' 		End If
			' 	End If
			' Next

			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_cars_exists_for_member(each_memb) = True Then
					cars_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					If STAT_INFORMATION(month_ind).stat_cars_one_exists(each_memb) = True Then
						Text 25, y_pos, 200, 10, "CARS   -   " & STAT_INFORMATION(month_ind).stat_cars_one_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_model(each_memb)
						Text 235, y_pos, 140, 10, "Use: " & STAT_INFORMATION(month_ind).stat_cars_one_use_info(each_memb)
						Text 385, y_pos, 85, 10, "HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_one_hc_clt_benefit_yn(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 110, 10, "Value: Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_one_trade_in_value(each_memb)
						Text 180, y_pos, 80, 10, "Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_one_loan_value(each_memb)
						Text 280, y_pos, 135, 10, "Value Source: " & STAT_INFORMATION(month_ind).stat_cars_one_value_source_info(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 135, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_cars_one_own_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_cars_one_own_verif_code(each_memb) = "N" Then Text 280, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 15
					End If

					If STAT_INFORMATION(month_ind).stat_cars_two_exists(each_memb) = True Then
						Text 25, y_pos, 200, 10, "CARS   -   " & STAT_INFORMATION(month_ind).stat_cars_two_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_model(each_memb)
						Text 235, y_pos, 140, 10, "Use: " & STAT_INFORMATION(month_ind).stat_cars_two_use_info(each_memb)
						Text 385, y_pos, 85, 10, "HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_two_hc_clt_benefit_yn(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 110, 10, "Value: Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_two_trade_in_value(each_memb)
						Text 180, y_pos, 80, 10, "Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_two_loan_value(each_memb)
						Text 280, y_pos, 135, 10, "Value Source: " & STAT_INFORMATION(month_ind).stat_cars_two_value_source_info(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 135, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_cars_two_own_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_cars_two_own_verif_code(each_memb) = "N" Then Text 280, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 15
					End If

					If STAT_INFORMATION(month_ind).stat_cars_three_exists(each_memb) = True Then
						Text 25, y_pos, 200, 10, "CARS   -   " & STAT_INFORMATION(month_ind).stat_cars_three_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_model(each_memb)
						Text 235, y_pos, 140, 10, "Use: " & STAT_INFORMATION(month_ind).stat_cars_three_use_info(each_memb)
						Text 385, y_pos, 85, 10, "HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_three_hc_clt_benefit_yn(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 110, 10, "Value: Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_three_trade_in_value(each_memb)
						Text 180, y_pos, 80, 10, "Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_three_loan_value(each_memb)
						Text 280, y_pos, 135, 10, "Value Source: " & STAT_INFORMATION(month_ind).stat_cars_three_value_source_info(each_memb)
						y_pos = y_pos + 10
						Text 60, y_pos, 135, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_cars_three_own_verif_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_cars_three_own_verif_code(each_memb) = "N" Then Text 280, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 15
					End If

				End If
			Next
			If y_pos = 25 Then
				Text 20, y_pos, 350, 10, "NO CARS panels have been entered in the csae file for the selected members."
				y_pos = y_pos + 10
				Text 50, y_pos, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				y_pos = y_pos + 15
				Text 20, y_pos, 350, 20, "If there are vehicle assets that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				y_pos = y_pos + 20
			End If

			Text 25, y_pos+5, 50, 10, "Vehicle Notes:"
			EditBox 75, y_pos, 395, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_cars_notes)
			y_pos = y_pos + 25

			start_y_pos = y_pos
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_rest_exists_for_member(each_memb) = True Then
					rest_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					If STAT_INFORMATION(month_ind).stat_rest_one_exists(each_memb) = True Then
						Text 25, y_pos, 135, 10, "REST   -   " & STAT_INFORMATION(month_ind).stat_rest_one_type_info(each_memb)
						Text 185, y_pos, 130, 10, "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_property_ownership_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_rest_one_ownership_verif_code(each_memb) = "NO" Then Text 320, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
						Text 58, y_pos, 130, 10, "Property Status: " & STAT_INFORMATION(month_ind).stat_rest_one_property_status_info(each_memb)
						Text 195, y_pos, 105, 10, " Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_one_market_value(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_value_verif_info(each_memb)
						y_pos = y_pos + 10
						Text 195, y_pos, 100, 10, "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_one_amount_owed(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_owed_verif_info(each_memb)
						y_pos = y_pos + 15
					End If
					If STAT_INFORMATION(month_ind).stat_rest_two_exists(each_memb) = True Then
						Text 25, y_pos, 135, 10, "REST   -   " & STAT_INFORMATION(month_ind).stat_rest_two_type_info(each_memb)
						Text 185, y_pos, 130, 10, "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_property_ownership_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_rest_two_ownership_verif_code(each_memb) = "NO" Then Text 320, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
						Text 58, y_pos, 130, 10, "Property Status: " & STAT_INFORMATION(month_ind).stat_rest_two_property_status_info(each_memb)
						Text 195, y_pos, 105, 10, " Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_two_market_value(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_value_verif_info(each_memb)
						y_pos = y_pos + 10
						Text 195, y_pos, 100, 10, "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_two_amount_owed(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_owed_verif_info(each_memb)
						y_pos = y_pos + 15
					End If
					If STAT_INFORMATION(month_ind).stat_rest_three_exists(each_memb) = True Then
						Text 25, y_pos, 135, 10, "REST   -   " & STAT_INFORMATION(month_ind).stat_rest_three_type_info(each_memb)
						Text 185, y_pos, 130, 10, "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_property_ownership_info(each_memb)
						If STAT_INFORMATION(month_ind).stat_rest_three_ownership_verif_code(each_memb) = "NO" Then Text 320, y_pos, 155, 10, "ADDED TO LIST OF VERIFICATIONS NEEDED"
						y_pos = y_pos + 10
						Text 58, y_pos, 130, 10, "Property Status: " & STAT_INFORMATION(month_ind).stat_rest_three_property_status_info(each_memb)
						Text 195, y_pos, 105, 10, " Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_three_market_value(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_value_verif_info(each_memb)
						y_pos = y_pos + 10
						Text 195, y_pos, 100, 10, "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_three_amount_owed(each_memb)
						Text 315, y_pos, 150, 10, "Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_owed_verif_info(each_memb)
						y_pos = y_pos + 15
					End If
				End If
			Next

			If y_pos = start_y_pos Then
				Text 20, y_pos, 350, 10, "NO REST panels have been entered in the csae file for the selected members."
				y_pos = y_pos + 10
				Text 50, y_pos, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				y_pos = y_pos + 15
				Text 20, y_pos, 350, 20, "If there are real estate assets that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				y_pos = y_pos + 20
			End If

			Text 25, y_pos+5, 50, 10, "Property Notes:"
			EditBox 75, y_pos, 395, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_rest_notes)
			y_pos = y_pos + 25


			' grp_len = 10

			' If cars_exists <> rest_exists Then grp_len = grp_len + 75
			' If grp_len = 10 Then grp_len = 155

			grp_len = y_pos - 10
			' grp_len = grp_len + 25
			GroupBox 10, 10, 465, grp_len, "Vehicles and Real Estate"

			' If grp_len = 20 Then grp_len = 70
			' GroupBox 10, 10, 465, grp_len, "Expenses"

			Text 500, 92, 55, 13, "CARS/REST"
		ElseIf page_display = show_expenses_page Then															'Expenses Page

			pded_exists = False
			coex_exists = False
			dcex_exists = False
			'PDED - person
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_pded_exists(each_memb) = True Then
					pded_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					Text 25, y_pos, 135, 10, "PDED   -   Deductions from PDED Exist"
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) <> "" Then
						If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) = "1" Then Text 60, y_pos, 420, 10, "Eligbile for PICKLE Disregard"
						If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) = "2" Then Text 60, y_pos, 420, 10, "Potentially Eligbile for PICKLE Disregard"
						y_pos = y_pos + 10
						Text 75, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_amt(each_memb) & " PICKLE Disregard Amount"
						y_pos = y_pos + 10
						Text 75, y_pos, 400, 10, "Current RSDI $ " & STAT_INFORMATION(month_ind).stat_pded_pickle_curr_RSDI(each_memb) & " less Threshold RSDI $ " & STAT_INFORMATION(month_ind).stat_pded_pickle_threshold_RSDI(each_memb) & ". Based on Threshold Date: " & STAT_INFORMATION(month_ind).stat_pded_pickle_threshold_date(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_disa_widow_deducation_yn(each_memb) = "Y" Then
						Text 60, y_pos, 135, 10, "Disabled Widow/ers Deduction applied"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_disa_adult_child_disregard_yn(each_memb) = "Y" Then
						Text 60, y_pos, 135, 10, "Disabled Adult Child Disregard applied"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_widow_deducation_yn(each_memb) = "Y" Then
						Text 60, y_pos, 135, 10, "Widow/ers Deduction applied"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_amt(each_memb) <> "" Then
						Text 60, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_amt(each_memb) & " Unearned Income Deduction Applied, Reason: " & STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_reason(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_amt(each_memb) <> "" Then
						Text 60, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_amt(each_memb) & " Earned Income Deduction Applied, Reason: " & STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_reason(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_extend_ma_epd_limits_yn(each_memb) = "Y" Then
						Text 60, y_pos, 135, 10, "MA-EPD Income/Asset Limits Extended"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_disa_student_child_disregard_yn(each_memb) = "Y" Then
						Text 60, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_disa_student_child_disregard_amt(each_memb) & " Blind/Disabled Student Child Disregard"
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_PASS_begin_date(each_memb) <> "" Then
						Text 60, y_pos, 420, 10, "PASS Plan   -   Begin Date: " & STAT_INFORMATION(month_ind).stat_pded_PASS_begin_date(each_memb) & " - End Date: " & STAT_INFORMATION(month_ind).stat_pded_PASS_end_date(each_memb)
						y_pos = y_pos + 10
						If STAT_INFORMATION(month_ind).stat_pded_PASS_earned_excluded(each_memb) <> "" Then
							Text 75, y_pos, 400, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_earned_excluded(each_memb) & " - Earned Income Excluded"
							y_pos = y_pos + 10
						End if
						If STAT_INFORMATION(month_ind).stat_pded_PASS_unea_excluded(each_memb) <> "" Then
							Text 75, y_pos, 400, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_unea_excluded(each_memb) & " - Unearned Income Excluded"
							y_pos = y_pos + 10
						End if
						If STAT_INFORMATION(month_ind).stat_pded_PASS_assets_excluded(each_memb) <> "" Then
							Text 75, y_pos, 400, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_assets_excluded(each_memb) & " - Assets Excluded"
							y_pos = y_pos + 10
						End if
					End If
					If STAT_INFORMATION(month_ind).stat_pded_guardianship_fee(each_memb) <> "" Then
						Text 60, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_guardianship_fee(each_memb) & " Guardianship Fee reduced from income."
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_pded_rep_payee_fee(each_memb) <> "" Then
						Text 60, y_pos, 420, 10, "$ " & STAT_INFORMATION(month_ind).stat_pded_rep_payee_fee(each_memb) & " Rep Payee Fee reduced from income."
						y_pos = y_pos + 10
					End If
				End If
			Next

			'COEX - person
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_coex_exists(each_memb) = True Then
					coex_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					Text 25, y_pos, 135, 10, "COEX   -   Court Ordered Expenses"
					y_pos = y_pos + 10
					Text 60, y_pos, 135, 10, "$ " & STAT_INFORMATION(month_ind).stat_coex_total_prosp_amt(each_memb) & " TOTAL Expense"
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_coex_support_prosp_amt(each_memb) <> "" Then
						Text 75, y_pos, 250, 10, "$ " & STAT_INFORMATION(month_ind).stat_coex_support_prosp_amt(each_memb) & " SUPPORT Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_support_verif_info(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_coex_alimony_prosp_amt(each_memb) <> "" Then
						Text 75, y_pos, 250, 10, "$ " & STAT_INFORMATION(month_ind).stat_coex_alimony_prosp_amt(each_memb) & " ALIMONY Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_alimony_verif_info(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_coex_tax_dep_prosp_amt(each_memb) <> "" Then
						Text 75, y_pos, 250, 10, "$ " & STAT_INFORMATION(month_ind).stat_coex_tax_dep_prosp_amt(each_memb) & " TAX DEPENDENT Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_tax_dep_verif_info(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_coex_other_prosp_amt(each_memb) <> "" Then
						Text 75, y_pos, 250, 10, "$ " & STAT_INFORMATION(month_ind).stat_coex_other_prosp_amt(each_memb) & " OTHER Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_other_verif_info(each_memb)
						y_pos = y_pos + 10
					End If
				End If
			Next

			'DCEX - person
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_dcex_exists(each_memb) = True Then
					dcex_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					Text 25, y_pos, 135, 10, "DCEX   -   Dependent Care Expenses"
					Text 150, y_pos, 300, 10, "Provider: " &  STAT_INFORMATION(month_ind).stat_dcex_provider(each_memb) & "   -   Reason: " & STAT_INFORMATION(month_ind).stat_dcex_reason_info(each_memb)
					y_pos = y_pos + 10
					If InStr(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb), ",") <> 0 Then
						dcex_child_array = split(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb), ",")
						dcex_amount_array = split(STAT_INFORMATION(month_ind).stat_dcex_prosp_amt_list(each_memb), ",")
						dcex_verif_array = split(STAT_INFORMATION(month_ind).stat_dcex_verif_info_list(each_memb), ",")
					Else
						dcex_child_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb))
						dcex_amount_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_prosp_amt_list(each_memb))
						dcex_verif_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_verif_info_list(each_memb))
					End If
					For dcex_child = 0 to UBound(dcex_child_array)
						Text 60, y_pos, 135, 10, "$ " & dcex_amount_array(dcex_child) & " for MEMB " & dcex_child_array(dcex_child) & ", Verif: " & dcex_verif_array(dcex_child)
						y_pos = y_pos + 10
					Next
					dcex_child_array = ""
					dcex_amount_array = ""
					dcex_verif_array = ""
				End If
			Next

			If pded_exists = False or coex_exists = False or dcex_exists = False Then
				y_pos = y_pos + 10
				If pded_exists = False Then panels_that_do_not_exists = panels_that_do_not_exists & "/PDED"
				If coex_exists = False Then panels_that_do_not_exists = panels_that_do_not_exists & "/COEX"
				If dcex_exists = False Then panels_that_do_not_exists = panels_that_do_not_exists & "/DCEX"
				If left(panels_that_do_not_exists, 1) = "/" Then panels_that_do_not_exists = right(panels_that_do_not_exists, len(panels_that_do_not_exists)-1)
				Text 20, y_pos, 300, 10, "This case does not have any " & panels_that_do_not_exists & " panels."
				y_pos = y_pos + 10
			End if

			grp_len = y_pos
			If grp_len = 20 Then grp_len = 70
			grp_len = grp_len + 25
			GroupBox 10, 10, 465, grp_len, "Expenses"

			If y_pos = 25 Then
				Text 20, 25, 350, 10, "NO PDED/COEX/DCEX panels have been entered in the csae file for the selected members."
				Text 50, 35, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				Text 20, 50, 350, 20, "If there are expenses that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."

				Text 20, 70, 350, 10, "CASE/NOTE will indicate NO EXPENSES, add any notes here that are relevant:"
				y_pos = 80
			Else
				y_pos = y_pos + 5
				Text 20, y_pos, 350, 10, "NOTES about Expenses and Deductions/Disregards:"
				y_pos = y_pos + 10
			End If
			EditBox 20, y_pos, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_expenses_general_notes)

			Text 505, 107, 55, 13, "Expenses"

		ElseIf page_display = show_other_page Then															'Other details Page
			acci_exists = False
			insa_exists = False
			faci_exists = False
			'TODO - each panel here should get it's own notes field

			y_pos = 25
			'ACCI - person
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_acci_exists(each_memb) = True Then
					acci_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 15
					Text 25, y_pos, 200, 10, "ACCI   -   Accident from " & STAT_INFORMATION(month_ind).stat_acci_injury_date(each_memb) & ". Medical cooperation: " & STAT_INFORMATION(month_ind).stat_acci_med_coop_yn(each_memb)
					y_pos = y_pos + 10
					Text 60, y_pos, 400, 10, "Accident Type: " & STAT_INFORMATION(month_ind).stat_acci_type_info(each_memb) & ". Involving MEMBS " & STAT_INFORMATION(month_ind).stat_acci_ref_numbers_list(each_memb)
					y_pos = y_pos + 10

					' y_pos = y_pos + 5
				End If
			Next
			If acci_exists = True Then
				Text 25, y_pos+5, 50, 10, "ACCI Notes:"
				EditBox 75, y_pos, 385, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_acci_notes)
				y_pos = y_pos + 20
			End If
			If acci_exists = False Then
				Text 20, y_pos, 205, 10, "NO ACCI Panel for any Member"
				y_pos = y_pos + 15
			End If

			'INSA - case
			For each_panel = 0 to UBound(STAT_INFORMATION(month_ind).stat_insa_exists)
				If STAT_INFORMATION(month_ind).stat_insa_exists(each_panel) = True Then
					insa_exists = True
					' Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					Text 20, y_pos, 350, 10, "INSA   -   Other Health Insurance through " & STAT_INFORMATION(month_ind).stat_insa_insurance_co(each_panel)
					y_pos = y_pos + 10
					Text 25, y_pos, 250, 10, "Covered MEMBS: " & STAT_INFORMATION(month_ind).stat_insa_covered_pers_list(each_panel)
					y_pos = y_pos + 10
					Text 25, y_pos, 350, 10, "Cooperation with OHI Requirements: " & STAT_INFORMATION(month_ind).stat_insa_coop_OHI_yn(each_panel) & "   -   Cooperation with CEHI Requirements: " & STAT_INFORMATION(month_ind).stat_insa_coop_cost_effective_yn(each_panel)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_insa_good_cause_code(each_panel) <> "_" And STAT_INFORMATION(month_ind).stat_insa_good_cause_code(each_panel) <> "N" Then
						Text 60, y_pos, 350, 10, STAT_INFORMATION(month_ind).stat_insa_good_cause_info(each_panel) & " - Claim Date: " & STAT_INFORMATION(month_ind).stat_insa_good_cause_claim_date(each_panel) & " - Evidence: " & STAT_INFORMATION(month_ind).stat_insa_coop_cost_effective_yn(each_panel)
						y_pos = y_pos + 10
					End If
					' y_pos = y_pos + 5
					' y_pos = y_pos + 5
				End If
			Next
			If insa_exists = True Then
				Text 25, y_pos+5, 50, 10, "INSA Notes:"
				EditBox 75, y_pos, 385, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_insa_notes)
				y_pos = y_pos + 20
			End If
			If insa_exists = False Then
				Text 20, y_pos, 205, 10, "NO INSA Panel for any Member"
				y_pos = y_pos + 15
			End If

			'TODO add PBEN - person
			'TODO add HCMI - person

			'FACI - person
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_faci_exists(each_memb) = True and STAT_INFORMATION(month_ind).stat_faci_currently_in_facility(each_memb) = True Then
					faci_exists = True
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 10
					Text 25, y_pos, 250, 10, "FACI   -   Resident in a Facility. In Date: " & STAT_INFORMATION(month_ind).stat_faci_date_in(each_memb)
					y_pos = y_pos + 10
					Text 55, y_pos, 200, 10, "Facility Name: " & STAT_INFORMATION(month_ind).stat_faci_name(each_memb)
					Text 260, y_pos, 200, 10, "Facility Type: " & STAT_INFORMATION(month_ind).stat_faci_type_info(each_memb)
					y_pos = y_pos + 10
					If STAT_INFORMATION(month_ind).stat_faci_waiver_type_info(each_memb) <> "" Then
						Text 55, y_pos, 150, 10, "Facility Waiver Type: " & STAT_INFORMATION(month_ind).stat_faci_waiver_type_info(each_memb)
						y_pos = y_pos + 10
					End If
					If STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb) <> "" Then
						Text 55, y_pos, 150, 10, "LTC Ineligible Reason: " & STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb)
						y_pos = y_pos + 10
					End If

					If excluded_time_case = True Then
						Text 55, y_pos, 300, 10, "EXCLUDED TIME CASE   -   County of Financial Responsibility: " & county_of_financial_responsibility
						y_pos = y_pos + 10
					End If
					' y_pos = y_pos + 5

				End If
				'TODO - add Excluded Time detail from SPEC if in FACI
				'TODO - advise if GRH faci is open and if NO BILS panel exists or BILS panel exists with NO 27 SERV type are listed. This is to support remembering about Remedial Care
			Next
			If faci_exists = True Then
				Text 25, y_pos+5, 50, 10, "FACI Notes:"
				EditBox 75, y_pos, 385, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_faci_notes)
				y_pos = y_pos + 20
			End If
			If faci_exists = False Then
				Text 20, y_pos, 350, 10, "NO FACI Panel for any Member, or NO FACI panel indicating currently IN a facility."
				y_pos = y_pos + 15
			End If

			grp_len = y_pos-5
			' If grp_len = 20 Then grp_len = 100
			If acci_exists = False and insa_exists = False and faci_exists = False Then
				grp_len = grp_len + 75

				Text 20, y_pos, 350, 10, "NO ACCI/INSA/FACI panels have been entered in the csae file for the selected members."
				y_pos = y_pos + 10
				Text 50, y_pos, 350, 10, "Selected Members for this case: MEMB " & replace(List_of_HH_membs_to_include, " ", "/")
				y_pos = y_pos + 15
				Text 20, y_pos, 350, 20, "If there are details from these panels that should be included for these members, CANCEL the Script, UPDATE MAXIS, and then rerun this script."
				y_pos = y_pos + 20

				Text 20, y_pos, 350, 10, "CASE/NOTE will indicate not add any other details, add any notes here that are relevant:"
				y_pos = y_pos + 10
				' y_pos = 80
			Else
				grp_len = grp_len + 35
				y_pos = y_pos + 5
				Text 20, y_pos, 350, 10, "NOTES about Miscelaneous Case/Person Information:"
				y_pos = y_pos + 10
			End If

			EditBox 20, y_pos, 440, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_other_general_notes)
			GroupBox 10, 10, 465, grp_len, "Other Information from Panels: ACCI/INSA/FACI"

			Text 512, 122, 55, 13, "Other"

		ElseIf page_display = bils_page Then															'BILS Page - this page only displays if there is a BILS panel

			Text 20, 25, 250, 10, "This case has medical bills entered into the BILS panel."

			Text 20, 40, 250, 10, "Check any box next to the bill to include the detail in the CASE/NOTE."

			y_pos = 55
			For each_bil = 0 to UBound(BILS_ARRAY, 2)
				CheckBox 25, y_pos, 95, 10, "MEMB " & BILS_ARRAY(bils_ref_numb_const, each_bil) & " from " & BILS_ARRAY(bils_date_const, each_bil), BILS_ARRAY(bils_checkbox, each_bil)
				Text 120, y_pos, 375, 10, "Gross: $ " & BILS_ARRAY(bils_gross_amt_const, each_bil) & ",  Service: " & BILS_ARRAY(bils_service_info_const, each_bil) & ",  Type: " & BILS_ARRAY(bils_expense_type_info_const, each_bil) & ",  Verif: " & BILS_ARRAY(bils_verif_info_const, each_bil)
				y_pos = y_pos + 10
			Next
			y_pos = y_pos + 5
			Text 20, y_pos, 350, 10, "Additional NOTES about BILS:"
			EditBox 20, y_pos+10, 440, 15, bils_notes

			grp_len = y_pos + 25
			GroupBox 10, 10, 465, grp_len, "Medical Bills - BILS"
		ElseIf page_display = imig_page Then															'IMIG Page - this page displays only if there is a IMIG request
			'TODO - SPON
			y_pos = 25
			For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
				If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = True and HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
					Text 20, y_pos, 205, 10, "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
					y_pos = y_pos + 10
					Text 25, y_pos, 250, 10, "IMIG   -   Immigration information. This resident is a Non-Citizen. Alien ID: " & STAT_INFORMATION(month_ind).stat_imig_alien_id_number(each_memb)
					y_pos = y_pos + 10
					Text 60, y_pos, 200, 10, "Status: " & STAT_INFORMATION(month_ind).stat_imig_status_info(each_memb) & ", entry date: " & STAT_INFORMATION(month_ind).stat_imig_entry_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_code(each_memb) <> "24" AND STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_code(each_memb) <> "__" Then Text 275, y_pos, 250, 10, "LPR Adjusted from " & STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_info(each_memb) & " on " & STAT_INFORMATION(month_ind).stat_imig_status_verif_code(each_memb)
					y_pos = y_pos + 10
					Text 60, y_pos, 150, 10, "Verification: " & STAT_INFORMATION(month_ind).stat_imig_status_verif_info(each_memb)
					y_pos = y_pos + 10
					Text 60, y_pos, 150, 10, "Nationality: " & STAT_INFORMATION(month_ind).stat_imig_nationality_info(each_memb)
					y_pos = y_pos + 10
					Text 60, y_pos, 375, 10, "40 Social Security Cr: " & STAT_INFORMATION(month_ind).stat_imig_40_credits_yn(each_memb) & "   -   Battered Spouse/Child: " & STAT_INFORMATION(month_ind).stat_imig_battered_pers_yn(each_memb) & "   -   Military Status: " & STAT_INFORMATION(month_ind).stat_imig_military_info(each_memb)
					y_pos = y_pos + 10
					Text 25, y_pos+5, 50, 10, "IMIG Notes:"
					EditBox 60, y_pos, 385, 15, EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_imig_notes(each_memb))
					y_pos = y_pos + 25
				End If
			Next
			grp_len = y_pos-10
			GroupBox 10, 10, 465, grp_len, "Immigration Information"
		ElseIf page_display = retro_page Then															'RETRO Page - this page displays only if there is a RETRO request
			y_pos = 25
			For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
				If HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = True and HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
					Text 15, y_pos, 300, 10, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, hc_memb) & " retro request to " & HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb)
					y_pos = y_pos + 15
				End If
			Next
			Text 15, y_pos, 460, 10, "The CASE/NOTE will include details from all the other dialog pages, including income and assets."
			y_pos = y_pos + 10
			Text 15, y_pos, 460, 10, "Here you can document details specific to the RETRO Request."
			y_pos = y_pos + 10
			Text 15, y_pos, 460, 10, "Any information entered into the fields specifically about verifications will be added to verifs needed."
			y_pos = y_pos + 15
			Text 35, y_pos+5, 65, 10, "Income Information:"
			EditBox 100, y_pos, 360, 15, retro_income_detail
			y_pos = y_pos + 20
			Text 40, y_pos+5, 60, 10, "Asset Information:"
			EditBox 100, y_pos, 360, 15, retro_asset_detail
			y_pos = y_pos + 20
			Text 30, y_pos+5, 70, 10, "Expense Information:"
			EditBox 100, y_pos, 360, 15, retro_expense_detail
			y_pos = y_pos + 20

			Groupbox 20, y_pos, 450, 70, "Verifs Needed:"
			y_pos = y_pos + 10
			Text 30, y_pos+5, 235, 10, "If Income verification from specific past month(s) is needed, list them here:"
			EditBox 265, y_pos, 195, 15, retro_income_verifs_months
			y_pos = y_pos + 20
			Text 30, y_pos+5, 230, 10, "If Asset verification from specific past month(s) is needed, list them here:"
			EditBox 260, y_pos, 200, 15, retro_asset_verifs_months
			y_pos = y_pos + 20
			Text 30, y_pos+5, 270, 10, "If Medical Expense verification from specific past month(s) is needed, list them here:"
			EditBox 300, y_pos, 160, 15, retro_expense_verifs_months
			y_pos = y_pos + 25
			Text 20, y_pos, 150, 10, "NOTES about RETRO Request:"
			EditBox 20, y_pos+10, 440, 15, retro_notes
			y_pos = y_pos + 35
			' Text 45, y_pos+5, 55, 10, "Something Else:"
			' EditBox 100, y_pos, 360, 15, edit_bot_info_5
			' y_pos = y_pos + 20
			' Text 100, y_pos, 200, 10, "If Income from specific past months is needed, list them here:"
			' EditBox 300, y_pos-5, 60, 15, retro_income_verifs_months
			' y_pos = y_pos + 20

			grp_len = y_pos-10
			GroupBox 10, 10, 465, grp_len, "RETRO Information"
		ElseIf page_display = ltc_page Then
			y_pos = 30
			ltc_info_in_stat = False
			For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
				If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
					For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
						If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) Then
							If HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) <> "" OR (STAT_INFORMATION(month_ind).stat_faci_exists(each_memb) = True and STAT_INFORMATION(month_ind).stat_faci_currently_in_facility(each_memb) = True) Then
								ltc_info_in_stat = True
								If HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb) <> "" Then
									Text 20, y_pos, 350, 10, "LTC Waiver: " & HEALTH_CARE_MEMBERS(DISA_waiver_info_const, hc_memb)
									y_pos = y_pos + 20
									Text 20, y_pos, 50, 10, "Waiver Notes:"
									EditBox 70, y_pos-5, 395, 15, HEALTH_CARE_MEMBERS(LTC_waiver_notes_const, hc_memb)
									y_pos = y_pos + 20
								End If
								If STAT_INFORMATION(month_ind).stat_faci_exists(each_memb) = True and STAT_INFORMATION(month_ind).stat_faci_currently_in_facility(each_memb) = True Then
									Text 20, y_pos, 350, 10, "Facility Info:  Name: " & STAT_INFORMATION(month_ind).stat_faci_name(each_memb)  & " - Type: " & STAT_INFORMATION(month_ind).stat_faci_type_info(each_memb)
									y_pos = y_pos + 10
									If STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb) <> "" Then
										Text 55, y_pos, 300, 10, "LTC Ineligible Reason: " & STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb)
										y_pos = y_pos + 10
									End If
									If excluded_time_case = True Then
										Text 55, y_pos, 300, 10, "EXCLUDED TIME CASE   -   County of Financial Responsibility: " & county_of_financial_responsibility
										y_pos = y_pos + 10
									End If
									y_pos = y_pos + 10
									Text 20, y_pos, 50, 10, "Facility Notes:"
									EditBox 70, y_pos-5, 395, 15, HEALTH_CARE_MEMBERS(LTC_facility_notes_const, hc_memb)
									y_pos = y_pos + 20
								End If
							End If
						End If
					Next
				End If
			Next
			If ltc_info_in_stat = False Then
				Text 20, y_pos, 400, 10, "NO Waiver or FACI information found for any Members you selected to process HC in this script run."
				y_pos = y_pos + 20
			End If
			Text 20, y_pos, 70, 10, "LTC Eligiblity Notes:"
			y_pos = y_pos + 10
			EditBox 20, y_pos, 445, 15, ltc_elig_notes
			y_pos = y_pos + 25
			Text 20, y_pos, 85, 10, "Information still needed:"
			y_pos = y_pos + 10
			EditBox 20, y_pos, 445, 15, ltc_info_still_needed
			y_pos = y_pos + 25

			grp_len = y_pos-10
			GroupBox 10, 10, 465, grp_len, "LTC Details"
		ElseIf page_display = last_page Then															'Final detail Page
			y_pos = 10
			If arep_name <> "" Then
				y_pos = y_pos + 10
				Text 20, y_pos, 150, 10, "AREP Information"
				' y_pos = y_pos + 10
				Text 275, y_pos, 150, 10, "Name: " & arep_name
				y_pos = y_pos + 10
				Text 25, y_pos, 300, 10, "Address: " & arep_addr_street & " " & arep_addr_city & ", " & arep_addr_state & " " & arep_addr_zip
				Text 275, y_pos, 75, 10, "Notices to AREP: " & forms_to_arep
				y_pos = y_pos + 10
				Text 25, y_pos, 150, 10, "Phone: " & arep_phone_one
				If arep_ext_one <> "" Then Text 175, y_pos, 75, 10, "Ext: " & arep_ext_one
				Text 275, y_pos, 75, 10, "MMIS Mail to AREP: " & mmis_mail_to_arep
				y_pos = y_pos + 5
			End If
			grp_len = y_pos
			GroupBox 10, 10, 465, grp_len, "AREP"
			If arep_name = "" Then Text 100, y_pos, 150, 10, "No AREP listed in this case."
			y_pos = y_pos + 15

			grp_pos = y_pos
			If swkr_name <> "" Then
				y_pos = y_pos + 10
				Text 20, y_pos, 150, 10, "SWKR Information"
				' y_pos = y_pos + 10
				Text 275, y_pos, 150, 10, "Name: " & swkr_name
				y_pos = y_pos + 10
				Text 25, y_pos, 300, 10, "Address: " & swkr_addr_street & " " & swkr_addr_city & ", " & swkr_addr_state & " " & swkr_addr_zip
				y_pos = y_pos + 10
				Text 25, y_pos, 150, 10, "Phone: " & swkr_phone_one
				If swkr_ext_one <> "" Then Text 175, y_pos, 75, 10, "Ext: " & swkr_ext_one
				Text 275, y_pos, 75, 10, "Notices to SWKR: " & notices_to_swkr_yn
				y_pos = y_pos + 5
			End If
			grp_len = y_pos - grp_pos + 10
			GroupBox 10, grp_pos, 465, grp_len, "SWKR"
			If swkr_name = "" Then Text 100, y_pos, 150, 10, "No SWKR listed in this case."
			y_pos = y_pos + 15

			grp_pos = y_pos
			y_pos = y_pos + 15

			Text 95, y_pos, 175, 10, "Has the Application been correctly Signed and Dated?"
			DropListBox 270, y_pos-5, 200, 15, "Select One..."+chr(9)+"Yes - All required signatures are on the application"+chr(9)+"No - Some applications or dates are missing", app_sig_status
			y_pos = y_pos + 20
			Text 170, y_pos, 100, 10, "If not, describe what is missing: "
			EditBox 270, y_pos-5, 200, 15, app_sig_notes
			y_pos = y_pos + 15

			For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
				If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
					GroupBox 20, y_pos, 445, 65, "MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, the_memb)
					y_pos = y_pos + 10
					Text 30, y_pos, 150, 10, "HC Eval Process:   " & HEALTH_CARE_MEMBERS(HC_eval_process_const, the_memb)
					Text 190, y_pos, 150, 10, "MA BASIS: " & HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, the_memb)
					Text 335, y_pos, 125, 10, "MSP BASIS: " & HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, the_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 75, 10, "Health Care Eval: "
					DropListBox 95, y_pos-5, 200, 15, "Select One..."+chr(9)+"Incomplete - need additional verificaitons"+chr(9)+"Incomplete - unclear information"+chr(9)+"Incomplete - other"+chr(9)+"Complete"+chr(9)+"More Processing Needed"+chr(9)+"Appears Ineligible", HEALTH_CARE_MEMBERS(hc_eval_status, the_memb)
					y_pos = y_pos + 20
					Text 30, y_pos, 70, 10, "Evaluation Notes:"
					EditBox 95, y_pos-5, 365, 15, HEALTH_CARE_MEMBERS(hc_eval_notes, the_memb)
					y_pos = y_pos + 15
				End if
			Next

			grp_len = y_pos - grp_pos + 10
			GroupBox 10, grp_pos, 465, grp_len, "Signatures and Status"

			' y_pos = y_pos + 15
			y_pos = y_pos + 15
			Text 15, y_pos, 150, 10, "Additional Case Details:"
			EditBox 15, y_pos+10, 530, 15, case_details_notes
			y_pos = y_pos + 30

			CheckBox 15, y_pos, 290, 10, "Check here to have the script update PND2 to show client delay (pending cases only).", client_delay_check
			y_pos = y_pos + 15
			If HC_form_name = "Breast and Cervical Cancer Coverage Group (DHS-3525)" Then
				CheckBox 15, y_pos, 350, 10, "Check here to have the script create a TIKL for 45 days (" & ma_bc_tikl_date & ") before REVW in " & revw_mm & "/" & revw_yy &".", MA_BC_end_of_cert_TIKL_check
			Else
				CheckBox 15, y_pos, 245, 10, "Check here to have the script create a TIKL to deny at the 45 day mark.", TIKL_check
			End If
		ElseIf page_display = verifs_page Then															'Verifs Page - this page displays only if there are verifs
			EditBox 700, 700, 50, 15, invisible_edit_box			'this is here to capture the attention of the tab order so people don't accidentally clear their verifs
			y_pos = 25
			Text 20, y_pos, 150, 10, "Verifications listed:"
			If verif_req_form_sent_date <> "" Then Text 200, y_pos, 150, 10, "Verification Request form Sent on " & verif_req_form_sent_date
			y_pos = y_pos + 10

			verifs_array = NULL
			verif_counter = 1
			verifs_needed = trim(verifs_needed)
			If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
			If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
			If InStr(verifs_needed, ";") <> 0 Then
				verifs_array = split(verifs_needed, ";")
			Else
				verifs_array = array(verifs_needed)
			End If

			For each verif_item in verifs_array
				verif_item = trim(verif_item)
				If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
				verif_counter = verif_counter + 1
				' Call write_variable_with_indent_in_CASE_NOTE(verif_item)
				Text 25, y_pos, 440, 10, verif_item
				y_pos = y_pos + 10
			Next
			y_pos = y_pos + 5
			Text 20, y_pos, 300, 10, "(Verifications to be added to CASE/NOTE)"
			y_pos = y_pos + 10

			grp_len = y_pos
			GroupBox 10, 10, 465, grp_len, "Verifications"
			y_pos = y_pos + 10

			Text 250, y_pos, 220, 20, "Pressing this button will remove all verifications from the list. You will need to press 'Update Verifications' to add verif information back."
			y_pos = y_pos + 20
			PushButton 350, y_pos, 120, 15, "Clear Verifs List", clear_verifs_btn

		' ElseIf page_display =  Then

		End If
		'panels I don't know what to do with
		'TODO - FCFC/FCPL
		'TODO - TRAN

		'This section of the code displays the buttons on the side and bottom
		Text 485, 5, 75, 10, "---   DIALOGS   ---"
		Text 485, 17, 10, 10, "1"
		Text 485, 32, 10, 10, "2"
		Text 485, 47, 10, 10, "3"
		Text 485, 62, 10, 10, "4"
		Text 485, 77, 10, 10, "5"
		Text 485, 92, 10, 10, "6"
		Text 485, 107, 10, 10, "7"
		Text 485, 122, 10, 10, "8"
		If page_display <> show_member_page 		Then PushButton 495, 15, 55, 13, "HC MEMBs", hc_memb_btn
		If page_display <> show_jobs_page 			Then PushButton 495, 30, 55, 13, "JOBS Income", jobs_inc_btn
		If page_display <> show_busi_page 			Then PushButton 495, 45, 55, 13, "BUSI Income", busi_inc_btn
		If page_display <> show_unea_page 			Then PushButton 495, 60, 55, 13, "UNEA Income", unea_inc_btn
		If page_display <> show_asset_page 			Then PushButton 495, 75, 55, 13, "Assets", assets_btn
		If page_display <> show_cars_rest_page 		Then PushButton 495, 90, 55, 13, "CARS/REST", cars_rest_btn
		If page_display <> show_expenses_page 		Then PushButton 495, 105, 55, 13, "Expenses", expenses_btn
		If page_display <> show_other_page 			Then PushButton 495, 120, 55, 13, "Other", other_btn

		btn_pos = 135								'these buttons only appear sometimes
		btn_count = 9
		If bils_exist = True Then
			Text 485, btn_pos + 2, 10, 10, btn_count
			If page_display <> bils_page 	Then PushButton 495, btn_pos, 55, 13, "BILS", bils_btn
			If page_display =  bils_page 	Then Text 515, btn_pos+2, 55, 13, "BILS"
			btn_pos = btn_pos + 15
			btn_count = btn_count + 1
		End If
		If imig_exists = True Then
			Text 485, btn_pos + 2, 10, 10, btn_count
			If page_display <> imig_page 	Then PushButton 495, btn_pos, 55, 13, "IMIG", imig_btn
			If page_display = imig_page 	Then Text 515, btn_pos+2, 55, 13, "IMIG"
			btn_pos = btn_pos + 15
			btn_count = btn_count + 1
		End If
		If case_has_retro_request = True Then
			Text 485, btn_pos + 2, 10, 10, btn_count
			If page_display <> retro_page 	Then PushButton 495, btn_pos, 55, 13, "RETRO", retro_btn
			If page_display = retro_page 	Then Text 510, btn_pos+2, 55, 13, "RETRO"
			btn_pos = btn_pos + 15
			btn_count = btn_count + 1
		End If
		If verifs_needed <> "" Then
			Text 485, btn_pos + 2, 10, 10, btn_count
			If page_display <> verifs_page 	Then PushButton 495, btn_pos, 55, 13, "Verifications",verifs_page_btn
			If page_display =  verifs_page 	Then Text 500, btn_pos+2, 55, 13, "Verifications"
			btn_pos = btn_pos + 15
			btn_count = btn_count + 1
		End If
		If ltc_waiver_request_yn = "Yes" Then
			Text 485, btn_pos + 2, 10, 10, btn_count
			If page_display <> ltc_page 	Then PushButton 495, btn_pos, 55, 13, "LTC Details", ltc_page_btn
			If page_display =  ltc_page 	Then Text 500, btn_pos+2, 55, 13, "LTC Details"
			btn_pos = btn_pos + 15
			btn_count = btn_count + 1
		End If

		Text 485, btn_pos + 2, 10, 10, btn_count
		last_page_numb = btn_count
		If page_display <> last_page 	Then PushButton 495, btn_pos, 55, 13, "App Info", last_btn
		If page_display =  last_page 	Then Text 505, btn_pos+2, 55, 13, "App Info"

		PushButton 20, 365, 130, 15, "Update Verifications", verif_button
		If verifs_needed <> "" Then Text 160, 368, 290, 10, "VERIFICATIONS EXIST"
		If page_display <> last_page Then PushButton 345, 365, 50, 15, "NEXT", next_btn
		PushButton 400, 365, 150, 15, "Complete Health Care Evaluation", completed_hc_eval_btn

	EndDialog
end function

function read_BILS_line(bil_row)
'This funciton reads a single BILS line and sets them to variables defined outside the dialog so they are not passed through
	EMReadScreen bil_ref_numb, 2, bil_row, 26
	EMReadScreen bil_date, 8, bil_row, 30
	EMReadScreen bil_serv_code, 2, bil_row, 40
	EMReadScreen bil_gross_amt, 9, bil_row, 45
	EMReadScreen bil_payments, 9, bil_row, 57
	EMReadScreen bil_verif_code, 2, bil_row, 67
	EMReadScreen bil_expense_type_code, 1, bil_row, 71
	EMReadScreen bil_old_priority, 2, bil_row, 75
	EMReadScreen bil_dependent_indicator, 1, bil_row, 79

	bil_date = replace(bil_date, " ", "/")

	If bil_serv_code = "" Then bil_serv_info = ""
	If bil_serv_code = "01" Then bil_serv_info = "Health Professional"
	If bil_serv_code = "03" Then bil_serv_info = "Surgery"
	If bil_serv_code = "04" Then bil_serv_info = "Chiropractic"
	If bil_serv_code = "05" Then bil_serv_info = "Maternity and Reproductive"
	If bil_serv_code = "07" Then bil_serv_info = "Hearing"
	If bil_serv_code = "08" Then bil_serv_info = "Vision"
	If bil_serv_code = "09" Then bil_serv_info = "Hospital"
	If bil_serv_code = "11" Then bil_serv_info = "Hospice"
	If bil_serv_code = "13" Then bil_serv_info = "SNF"
	If bil_serv_code = "14" Then bil_serv_info = "Dental"
	If bil_serv_code = "15" Then bil_serv_info = "Rx Drug/Non-Durable Supply"
	If bil_serv_code = "16" Then bil_serv_info = "Home Health"
	If bil_serv_code = "17" Then bil_serv_info = "Diagnostic"
	If bil_serv_code = "18" Then bil_serv_info = "Mental Health"
	If bil_serv_code = "19" Then bil_serv_info = "Rehabilitation Habilitation"
	If bil_serv_code = "21" Then bil_serv_info = "Durable Med Equip/Supplies"
	If bil_serv_code = "22" Then bil_serv_info = "Medical Trans"
	If bil_serv_code = "24" Then bil_serv_info = "Waivered Serv"
	If bil_serv_code = "25" Then bil_serv_info = "Medicare Prem"
	If bil_serv_code = "26" Then bil_serv_info = "Dental or Health Prem"
	If bil_serv_code = "27" Then bil_serv_info = "Remedial Care"
	If bil_serv_code = "28" Then bil_serv_info = "Non-FFP MCRE Service"
	If bil_serv_code = "30" Then bil_serv_info = "Alternative Care"
	If bil_serv_code = "31" Then bil_serv_info = "MCSHN"
	If bil_serv_code = "32" Then bil_serv_info = "Ins Ext Prog"
	If bil_serv_code = "34" Then bil_serv_info = "CW-TCM"
	If bil_serv_code = "37" Then bil_serv_info = "Pay-In Spdn"
	If bil_serv_code = "42" Then bil_serv_info = "Access Services"
	If bil_serv_code = "43" Then bil_serv_info = "Chemical Dep"
	If bil_serv_code = "44" Then bil_serv_info = "Nutritional Services"
	If bil_serv_code = "45" Then bil_serv_info = "Organ/Tissue Transplant"
	If bil_serv_code = "46" Then bil_serv_info = "Out-Of-Area Services"
	If bil_serv_code = "47" Then bil_serv_info = "Copayment/Deductible"
	If bil_serv_code = "49" Then bil_serv_info = "Preventative Care"
	If bil_serv_code = "99" Then bil_serv_info = "Other"

	If bil_verif_code = "__" Then bil_verif_info = ""
	If bil_verif_code = "01" Then bil_verif_info = "Billing Statement"
	If bil_verif_code = "02" Then bil_verif_info = "Explanation of Benefit"
	If bil_verif_code = "03" Then bil_verif_info = "Client Statment"
	If bil_verif_code = "04" Then bil_verif_info = "Credit/Loan Statement"
	If bil_verif_code = "05" Then bil_verif_info = "Provider Statement"
	If bil_verif_code = "06" Then bil_verif_info = "Other"
	If bil_verif_code = "NO" Then bil_verif_info = "No Verif provided"

	If bil_expense_type_code = "_" Then bil_expense_type_info = "Unknown"
	If bil_expense_type_code = "H" Then bil_expense_type_info = "Health Ins, Other Premium"
	If bil_expense_type_code = "P" Then bil_expense_type_info = "Not Covered, Non-Reimbursed"
	If bil_expense_type_code = "M" Then bil_expense_type_info = "Old, Unpaid Medical Bills"
	If bil_expense_type_code = "R" Then bil_expense_type_info = "Reimbursable"
end function

function read_person_based_STAT_info()
'reading additional information from STAT for each person requesting HC
'This is seperate from the STAT Class because it is specific to those requesting HC
	Call navigate_to_MAXIS_screen("STAT", "DISA")													'reading DISA information
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

	Call navigate_to_MAXIS_screen("STAT", "PREG")														'reading information from PREG
	Call write_value_and_transmit(HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb), 20, 76)
	EMReadScreen preg_version, 1, 2, 78
	If preg_version = "1" Then
		HEALTH_CARE_MEMBERS(PREG_exists_const, hc_memb) = True
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_due_date_const, hc_memb), 8, 10, 53
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, hc_memb), 1, 6, 75
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_end_date_const, hc_memb), 8, 12, 53
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, hc_memb), 1, 12, 75
		EMReadScreen HEALTH_CARE_MEMBERS(PREG_multiple_const, hc_memb), 1, 14, 53

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
		HEALTH_CARE_MEMBERS(PREG_multiple_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(PREG_multiple_const, hc_memb), "_", "")
	End If


	Call navigate_to_MAXIS_screen("STAT", "PARE")														'reading information from PARE
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

	Call navigate_to_MAXIS_screen("STAT", "MEDI")												'reading information from MEDI
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
'this function is specific to this script and will use the ButtonPressed information to select which page to display in the dialog
	If ButtonPressed = -1 Then				'If the user presses the 'Enter' key, it is the same as the 'OK Button' even if it is not on the dialog
		If page_display = last_page Then ButtonPressed = completed_hc_eval_btn			'If 'Enter' is pressed, and we are on the last page, it is like pressing the completed button
		If page_display <> last_page Then ButtonPressed = next_btn						'if 'Enter' is pressed on any other page, it is like pressing the 'Next' button
	End if
	For i = 0 to Ubound(HEALTH_CARE_MEMBERS, 2)											'Looking for if a member button is pressed on page one.
		If ButtonPressed = HEALTH_CARE_MEMBERS(pers_btn_one_const, i) Then
			If page_display = show_member_page Then selected_memb = i					'if a button is pressed for a member, it sets the member index to the index on that button
		End If
	Next
	If ButtonPressed = next_btn Then page_display = page_display + 1					'incrementing the page to display as these are numeric
	If page_display = bils_page and bils_exist = False Then page_display = page_display + 1
	If page_display = imig_page and imig_exists = False Then page_display = page_display + 1
	If page_display = retro_page and case_has_retro_request = False Then page_display = page_display + 1
	If page_display = verifs_page and verifs_needed = "" Then page_display = page_display + 1
	If page_display = ltc_page and ltc_waiver_request_yn <> "Yes" Then page_display = page_display + 1
	If page_display > last_btn Then page_display = last_page							'making sure we don't go above the last page

	'Each button pressed sets page_dsiplay to the page associated with the button
	If ButtonPressed = hc_memb_btn Then page_display = show_member_page
	If ButtonPressed = jobs_inc_btn Then page_display = show_jobs_page
	If ButtonPressed = busi_inc_btn Then page_display = show_busi_page
	If ButtonPressed = unea_inc_btn Then page_display = show_unea_page
	If ButtonPressed = assets_btn Then page_display = show_asset_page
	If ButtonPressed = cars_rest_btn Then page_display = show_cars_rest_page
	If ButtonPressed = expenses_btn Then page_display = show_expenses_page
	If ButtonPressed = other_btn Then page_display = show_other_page
	If ButtonPressed = bils_btn Then page_display = bils_page
	If ButtonPressed = imig_btn Then page_display = imig_page
	If ButtonPressed = retro_btn Then page_display = retro_page
	If ButtonPressed = verifs_page_btn Then page_display = verifs_page
	If ButtonPressed = ltc_page_btn Then page_display = ltc_page
	If ButtonPressed = last_btn Then page_display = last_page

	If ButtonPressed = clear_verifs_btn Then			'If the 'Clear Verifs' button is pressed, we blank out 'verifs_needed' and go to the last page
		verifs_needed = ""
		page_display = last_page
	End If

	'here we add some verif information if needed
	If right(verifs_needed, 1) = ";" Then verifs_needed = verifs_needed & " "
	If right(verifs_needed, 2) <> "; " Then verifs_needed = verifs_needed & "; "
	If verifs_needed = "; " Then verifs_needed = ""

	If ma_bc_authorization_form_missing_checkbox = checked and trim(ma_bc_authorization_form) <> "" Then
		If Instr(verifs_needed, "MA-BC treatment/screening form needed to process MA-BC eligibility.") = 0 Then
			verifs_needed = verifs_needed & "MA-BC treatment/screening form needed to process MA-BC eligibility.; "
		End If
	End If
	For each_hh_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
		If HEALTH_CARE_MEMBERS(MEDI_application_requred_checkbox_const, each_hh_memb) = checked Then
			If InStr(verifs_needed, "Application for MEDICARE required for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, each_hh_memb)) = 0 Then
				If HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) <> "" Then verifs_needed = verifs_needed & "Application for MEDICARE required for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, each_hh_memb) & ", referral date: " & HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) & "; "
				If HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) = "" Then verifs_needed = verifs_needed & "Application for MEDICARE required for MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, each_hh_memb) & "; "
			End If
		End If
	Next
	If avs_form_status = "No Form Received" and InStr(verifs_needed, "Signed AVS Authorization Form (DHS-7823)") = 0 Then verifs_needed = verifs_needed & "Signed AVS Authorization Form (DHS-7823); "
	If avs_form_status = "Incomplete - Received - Not Signed/Dated Correctly" and InStr(verifs_needed, "AVS Authorization Form (DHS-7823) needs to be signed/dated correctly.") = 0 Then verifs_needed = verifs_needed & "AVS Authorization Form (DHS-7823) needs to be signed/dated correctly.; "
	If avs_form_status = "Incomplete - Received - Not Signed by all Required Members" and InStr(verifs_needed, "AVS Authorization Form (DHS-7823) needs to be signed by all required persons.") = 0 Then verifs_needed = verifs_needed & "AVS Authorization Form (DHS-7823) needs to be signed by all required persons.; "
	retro_income_verifs_months = trim(retro_income_verifs_months)
	retro_asset_verifs_months = trim(retro_asset_verifs_months)
	retro_expense_verifs_months = trim(retro_expense_verifs_months)
	If retro_income_verifs_months <> "" AND InStr(verifs_needed, retro_income_verifs_months) = 0 Then verifs_needed = verifs_needed & "Retro Months Income Information (" & retro_income_verifs_months & "); "
	If retro_asset_verifs_months <> "" AND InStr(verifs_needed, retro_asset_verifs_months) = 0 Then verifs_needed = verifs_needed & "Retro Months Asset Information (" & retro_asset_verifs_months & "); "
	If retro_expense_verifs_months <> "" AND InStr(verifs_needed, retro_expense_verifs_months) = 0 Then verifs_needed = verifs_needed & "Retro Months Expense Information (" & retro_expense_verifs_months & "); "

	If ButtonPressed = completed_hc_eval_btn Then leave_loop = TRUE		'if the button to complete the review is pressed, the movement allows you to leave the loop
end function

function verification_dialog()
'this function is script specific to display a dialog allowing selection of verifications
    If ButtonPressed = verif_button Then
        Do
            verif_err_msg = ""

            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 610, 395, "Select Verifications"
              Text 280, 10, 120, 10, "Date Verification Request Form Sent:"
              EditBox 400, 5, 50, 15, verif_req_form_sent_date

              Groupbox 5, 35, 520, 130, "Personal and Household Information"

              CheckBox 10, 50, 75, 10, "Verification of ID for ", id_verif_checkbox
              ComboBox 90, 45, 150, 45, verification_memb_list, id_verif_memb
              CheckBox 300, 50, 100, 10, "Social Security Number for ", ssn_checkbox
              ComboBox 405, 45, 110, 45, verification_memb_list, ssn_verif_memb

              CheckBox 10, 70, 70, 10, "US Citizenship for ", us_cit_status_checkbox
              ComboBox 85, 65, 150, 45, verification_memb_list, us_cit_verif_memb
              CheckBox 300, 70, 85, 10, "Immigration Status for", imig_status_checkbox
              ComboBox 390, 65, 125, 45, verification_memb_list, imig_verif_memb

              CheckBox 10, 90, 90, 10, "Proof of relationship for ", relationship_checkbox
              ComboBox 105, 85, 150, 45, verification_memb_list, relationship_one_verif_memb
              Text 260, 90, 90, 10, "and"
              ComboBox 280, 85, 150, 45, verification_memb_list, relationship_two_verif_memb

              CheckBox 10, 110, 85, 10, "Student Information for ", student_info_checkbox
              ComboBox 100, 105, 150, 45, verification_memb_list, student_verif_memb
              Text 255, 110, 10, 10, "at"
              EditBox 270, 105, 150, 15, student_verif_source

              CheckBox 10, 130, 85, 10, "Proof of Pregnancy for", preg_checkbox
              ComboBox 100, 125, 150, 45, verification_memb_list, preg_verif_memb

              CheckBox 10, 150, 115, 10, "Illness/Incapacity/Disability for", illness_disability_checkbox
              ComboBox 130, 145, 150, 45, verification_memb_list, disa_verif_memb
              Text 285, 150, 30, 10, "verifying:"
              EditBox 320, 145, 150, 15, disa_verif_type

              GroupBox 5, 165, 520, 50, "Income Information"

              CheckBox 10, 180, 45, 10, "Income for ", income_checkbox
              ComboBox 60, 175, 140, 45, verification_memb_list, income_verif_memb
              Text 205, 180, 15, 10, "from"
              ComboBox 225, 175, 125, 45, income_source_list, income_verif_source
              Text 355, 180, 10, 10, "for"
              EditBox 370, 175, 145, 15, income_verif_time

              CheckBox 10, 200, 85, 10, "Employment Status for ", employment_status_checkbox
              ComboBox 100, 195, 150, 45, verification_memb_list, emp_status_verif_memb
              Text 255, 200, 10, 10, "at"
              ComboBox 270, 195, 150, 45, employment_source_list, emp_status_verif_source

              GroupBox 5, 215, 520, 50, "Expense Information"

              CheckBox 10, 230, 105, 10, "Educational Funds/Costs for", educational_funds_cost_checkbox
              ComboBox 120, 225, 150, 45, verification_memb_list, stin_verif_memb

              CheckBox 10, 250, 65, 10, "Shelter Costs for ", shelter_checkbox
              ComboBox 80, 245, 150, 45, verification_memb_list, shelter_verif_memb
              checkBox 240, 250, 175, 10, "Check here if this verif is NOT MANDATORY", shelter_not_mandatory_checkbox

              GroupBox 5, 265, 600, 30, "Asset Information"

              CheckBox 10, 280, 70, 10, "Bank Account for", bank_account_checkbox
              ComboBox 80, 275, 150, 45, verification_memb_list, bank_verif_memb
              Text 235, 280, 45, 10, "account type"
              ComboBox 285, 275, 145, 45, "Select or Type"+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Certificate of Deposit (CD)"+chr(9)+"Stock"+chr(9)+"Money Market", bank_verif_type
              Text 435, 280, 10, 10, "for"
              EditBox 450, 275, 150, 15, bank_verif_time

              Text 5, 305, 20, 10, "Other:"
              EditBox 30, 300, 570, 15, other_verifs
              Checkbox 10, 320, 200, 10, "Check here to have verifs numbered in the CASE/NOTE.", number_verifs_checkbox

              ButtonGroup ButtonPressed
                PushButton 485, 10, 50, 15, "FILL", fill_button
                PushButton 540, 10, 60, 15, "Return to Dialog", return_to_dialog_button
              Text 10, 340, 580, 50, verifs_needed
              Text 10, 10, 235, 10, "Check the boxes for any verification you want to add to the CASE/NOTE."
              Text 10, 20, 470, 10, "Note: After you press 'Fill' or 'Return to Dialog' the information from the boxes will fill in the Verification Field and the boxes will be 'unchecked'."
            EndDialog

            dialog Dialog1

            If ButtonPressed = 0 Then
                id_verif_checkbox = unchecked
                us_cit_status_checkbox = unchecked
                imig_status_checkbox = unchecked
                ssn_checkbox = unchecked
                relationship_checkbox = unchecked
                income_checkbox = unchecked
                employment_status_checkbox = unchecked
                student_info_checkbox = unchecked
                educational_funds_cost_checkbox = unchecked
                shelter_checkbox = unchecked
                bank_account_checkbox = unchecked
                preg_checkbox = unchecked
                illness_disability_checkbox = unchecked
            End If
            If ButtonPressed = -1 Then ButtonPressed = fill_button

			'verif dialog err msg handling
            If id_verif_checkbox = checked AND (id_verif_memb = "Select or Type Member" OR trim(id_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs ID verified."
            If us_cit_status_checkbox = checked AND (us_cit_verif_memb = "Select or Type Member" OR trim(us_cit_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs citizenship verified."
            If imig_status_checkbox = checked AND (imig_verif_memb = "Select or Type Member" OR trim(imig_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member that needs immigration status verified."
            If ssn_checkbox = checked AND (ssn_verif_memb = "Select or Type Member" OR trim(ssn_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need social security number."
            If relationship_checkbox = checked Then
                If relationship_one_verif_memb = "Select or Type Member" OR trim(relationship_one_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
                If relationship_two_verif_memb = "Select or Type Member" OR trim(relationship_two_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the two household members whose relationship needs to be verified."
            End If
            If income_checkbox = checked Then
                If income_verif_memb = "Select or Type Member" OR trim(income_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose income needs to be verified."
                If trim(income_verif_source) = "" OR trim(income_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of income to be verified."
                If trim(income_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the income verification needed."
            End If
            If employment_status_checkbox = checked Then
                If trim(emp_status_verif_source) = "" OR trim(emp_status_verif_source) = "Select or Type Source" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of the employment that needs status verified."
                If emp_status_verif_memb = "Select or Type Member" OR trim(emp_status_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose employment status needs to be verified."
            End If
            If student_info_checkbox = checked Then
                If trim(student_verif_source) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the source of school information to be verified"
                If student_verif_memb = "Select or Type Member" OR trim(student_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member for which we need school verification."
            End If
            If educational_funds_cost_checkbox = checked AND (stin_verif_memb = "Select or Type Member" OR trim(stin_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member with educational funds and costs we need verified."
            If shelter_checkbox = checked AND (shelter_verif_memb = "Select or Type Member" OR trim(shelter_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose shelter expense we need verified."
            If bank_account_checkbox = checked Then
                If trim(bank_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type of bank account to verify."
                If bank_verif_memb = "Select or Type Member" OR trim(bank_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose bank account we need verified."
                If trim(bank_verif_time) = "[Enter Time Frame]" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the time frame of the bank account verification needed."
            End If
            If preg_checkbox = checked AND (preg_verif_memb = "Select or Type Member" OR trim(preg_verif_memb) = "") Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose pregnancy needs to be verified."
            If illness_disability_checkbox = checked Then
                If trim(disa_verif_type) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Enter the type (or details) of the illness/incapacity/disability that need to be verified."
                If disa_verif_memb = "Select or Type Member" OR trim(disa_verif_memb) = "" Then verif_err_msg = verif_err_msg & vbNewLine & "* Indicate the household member whose illness/incapacity/disability needs to be verified."
            End If

            If verif_err_msg = "" Then
				'adding detail to verif information based on information entered into the verifs line
				If right(verifs_needed, 1) = ";" Then verifs_needed = verifs_needed & " "
				If right(verifs_needed, 2) <> "; " Then verifs_needed = verifs_needed & "; "
                If id_verif_checkbox = checked Then
                    If IsNumeric(left(id_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Identity for Memb " & id_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Identity for " & id_verif_memb & ".; "
                    End If
                    id_verif_checkbox = unchecked
                    id_verif_memb = ""
                End If
                If us_cit_status_checkbox = checked Then
                    If IsNumeric(left(us_cit_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "US Citizenship for Memb " & us_cit_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "US Citizenship for " & us_cit_verif_memb & ".; "
                    End If
                    us_cit_status_checkbox = unchecked
                    us_cit_verif_memb = ""
                End If
                If imig_status_checkbox = checked Then
                    If IsNumeric(left(imig_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Immigration documentation for Memb " & imig_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Immigration documentation for " & imig_verif_memb & ".; "
                    End If
                    imig_status_checkbox = unchecked
                    imig_verif_memb = ""
                End If
                If ssn_checkbox = checked Then
                    If IsNumeric(left(ssn_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Social Security number for Memb " & ssn_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Social Security number for " & ssn_verif_memb & ".; "
                    End If
                    ssn_checkbox = unchecked
                    ssn_verif_memb = ""
                End If
                If relationship_checkbox = checked Then
                    If IsNumeric(left(relationship_one_verif_memb, 2)) = TRUE AND IsNumeric(left(relationship_two_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Relationship between Memb " & relationship_one_verif_memb & " and Memb " & relationship_two_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Relationship between " & relationship_one_verif_memb & " and " & relationship_two_verif_memb & ".; "
                    End If
                    relationship_checkbox = unchecked
                    relationship_one_verif_memb = ""
                    relationship_two_verif_memb = ""
                End If
                If income_checkbox = checked Then
                    If IsNumeric(left(income_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Income for Memb " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    Else
                        verifs_needed = verifs_needed & "Income for " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    End If
                    income_checkbox = unchecked
                    income_verif_source = ""
                    income_verif_memb = ""
                    income_verif_time = ""
                End If
                If employment_status_checkbox = checked Then
                    If IsNumeric(left(emp_status_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Employment Status for Memb " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    Else
                        verifs_needed = verifs_needed & "Employment Status for " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    End If
                    employment_status_checkbox = unchecked
                    emp_status_verif_memb = ""
                    emp_status_verif_source = ""
                End If
                If student_info_checkbox = checked Then
                    If IsNumeric(left(student_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Student information for Memb " & student_verif_memb & " at " & student_verif_source & ".; "
                    Else
                        verifs_needed = verifs_needed & "Student information for " & student_verif_memb & " at " & student_verif_source & ".; "
                    End If
                    student_info_checkbox = unchecked
                    student_verif_memb = ""
                    student_verif_source = ""
                End If
                If educational_funds_cost_checkbox = checked Then
                    If IsNumeric(left(stin_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Educational funds and costs for Memb " & stin_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Educational funds and costs for " & stin_verif_memb & ".; "
                    End If
                    educational_funds_cost_checkbox = unchecked
                    stin_verif_memb = ""
                End If
                If shelter_checkbox = checked Then
                    If IsNumeric(left(shelter_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Shelter costs for Memb " & shelter_verif_memb & ". "
                    Else
                        verifs_needed = verifs_needed & "Shelter costs for " & shelter_verif_memb & ". "
                    End If
                    If shelter_not_mandatory_checkbox = checked Then verifs_needed = verifs_needed & " THIS VERIFICATION IS NOT MANDATORY."
                    verifs_needed = verifs_needed & "; "
                    shelter_checkbox = unchecked
                    shelter_verif_memb = ""
                End If
                If bank_account_checkbox = checked Then
                    If IsNumeric(left(bank_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & bank_verif_type & " account for Memb " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    Else
                        verifs_needed = verifs_needed & bank_verif_type & " account for " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    End If
                    bank_account_checkbox = unchecked
                    bank_verif_type = ""
                    bank_verif_memb = ""
                    bank_verif_time = ""
                End If
                If preg_checkbox = checked Then
                    If IsNumeric(left(preg_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Pregnancy for Memb " & preg_verif_memb & ".; "
                    Else
                        verifs_needed = verifs_needed & "Pregnancy for " & preg_verif_memb & ".; "
                    End If
                    preg_checkbox = unchecked
                    preg_verif_memb = ""
                End If
                If illness_disability_checkbox = checked Then
                    If IsNumeric(left(disa_verif_memb, 2)) = TRUE Then
                        verifs_needed = verifs_needed & "Ill/Incap or Disability for Memb " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    Else
                        verifs_needed = verifs_needed & "Ill/Incap or Disability for " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    End If
                    illness_disability_checkbox = unchecked
                    disa_verif_memb = ""
                    disa_verif_type = ""
                End If
                other_verifs = trim(other_verifs)
                If other_verifs <> "" Then verifs_needed = verifs_needed & other_verifs & "; "
                other_verifs = ""
				verifs_needed = trim(verifs_needed)
				If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed)-1)
            Else
                MsgBox "Additional detail about verifications to note is needed:" & vbNewLine & verif_err_msg
            End If

            If ButtonPressed = fill_button Then verif_err_msg = "LOOP" & verif_err_msg
        Loop until verif_err_msg = ""
        ButtonPressed = verif_button			'this takes us back to the verif display on the main dialog
    End If
end function

function write_header_and_detail_in_CASE_NOTE(header, variable)
'--- This function creates an indent for the header and then indents the detail after the header, this is specific to this script
'~~~~~ header: name of the field to update. Put header in "".
'~~~~~ variable: variable from script to be written into CASE note
'===== Keywords: MAXIS, header, CASE note
	If trim(variable) <> "" then
		EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
		noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
		'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
			If character_test <> "" or noting_row >= 18 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 18 then
					EMSendKey "<PF8>"
					EMWaitReady 0, 0

                    EMReadScreen check_we_went_to_next_page, 75, 24, 2
                    check_we_went_to_next_page = trim(check_we_went_to_next_page)

					'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
					EMReadScreen end_of_case_note_check, 1, 24, 2
					If end_of_case_note_check = "A" then
						EMSendKey "<PF3>"												'PF3s
						EMWaitReady 0, 0
						EMSendKey "<PF9>"												'PF9s (opens new note)
						EMWaitReady 0, 0
						EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
						EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
						noting_row = 5													'Resets this variable to work in the new locale
                    ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                        noting_row = 4
                        Do
                            EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                            character_test = trim(character_test)
                            If character_test <> "" then noting_row = noting_row + 1
                        Loop until character_test = ""
                    Else
						noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                    End If
				End if
			End if
		Loop until character_test = ""

		'Looks at the length of the header. This determines the indent for the rest of the info. Going with a maximum indent of 18.
		' If len(header) >= 14 then
		' 	indent_length = 18	'It's four more than the header text to account for the asterisk, the colon, and the spaces.
		' Else
		' 	indent_length = len(header) + 7 'It's four more for the reason explained above.
		' End if
		indent_length = 9
		'Writes the header
		EMWriteScreen "     " & header & ": ", noting_row, noting_col

		'Determines new noting_col based on length of the header length (header + 4 to account for asterisk, colon, and spaces).
		noting_col = noting_col + (len(header) + 7)

		'Splits the contents of the variable into an array of words
        variable = trim(variable)
        If right(variable, 1) = ";" Then variable = left(variable, len(variable) - 1)
        If left(variable, 1) = ";" Then variable = right(variable, len(variable) - 1)
		variable_array = split(variable, " ")

		For each word in variable_array
			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the header
			If len(word) + noting_col > 80 then
				noting_row = noting_row + 1
				noting_col = 3
			End if

			'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

                EMReadScreen check_we_went_to_next_page, 75, 24, 2
                check_we_went_to_next_page = trim(check_we_went_to_next_page)

                'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
                EMReadScreen end_of_case_note_check, 1, 24, 2
                If end_of_case_note_check = "A" then
                    EMSendKey "<PF3>"												'PF3s
                    EMWaitReady 0, 0
                    EMSendKey "<PF9>"												'PF9s (opens new note)
                    EMWaitReady 0, 0
                    EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
                    EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
                    noting_row = 5													'Resets this variable to work in the new locale
                ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
                    noting_row = 4
                    Do
                        EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
                        character_test = trim(character_test)
                        If character_test <> "" then noting_row = noting_row + 1
                    Loop until character_test = ""
                Else
                    noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
                End If
			End if

			'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
			If noting_col = 3 then
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
			If right(word, 1) = ";" then
				noting_row = noting_row + 1
				noting_col = 3
				EMWriteScreen space(indent_length), noting_row, noting_col
				noting_col = noting_col + indent_length
			End if

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
	End if
end function

'END FUNCTIONS BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================
'constants for the HEALTH_CARE_MEMBERS array
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
Const member_has_retro_request			= 31

Const show_hc_detail_const 				= 32
Const DISA_exists_const 				= 33
Const DISA_start_date_const 			= 34
Const DISA_end_date_const 				= 35
Const DISA_cert_start_const 			= 36
Const DISA_cert_end_const 				= 37
Const DISA_hc_status_code_const 		= 38
Const DISA_hc_status_info_const 		= 39
Const DISA_hc_verif_code_const 			= 40
Const DISA_hc_verif_info_const 			= 41
Const DISA_waiver_code_const			= 42
Const DISA_waiver_info_const			= 43
Const DISA_notes_const					= 44
Const PREG_exists_const 				= 45
Const PREG_due_date_const 				= 46
Const PREG_due_date_verif_const 		= 47
Const PREG_end_date_const 				= 48
Const PREG_end_date_verif_const 		= 49
Const PREG_multiple_const				= 50
Const PREG_notes_const					= 51
Const PARE_exists_const 				= 52
Const PARE_list_of_children_const 		= 53
Const PARE_notes_const					= 54
Const MEDI_exists_const 				= 55
Const MEDI_part_A_premium_const 		= 56
Const MEDI_part_B_premium_const 		= 57
Const MEDI_part_A_start_const 			= 58
Const MEDI_part_A_end_const 			= 59
Const MEDI_part_B_start_const 			= 60
Const MEDI_part_B_end_const 			= 61
Const MEDI_info_source_const 			= 62
Const MEDI_application_requred_checkbox_const	= 63
Const MEDI_referral_date_const 			= 64
Const MEDI_notes_const					= 65
Const HC_eval_process_const 			= 66
Const HC_basis_of_elig_const 			= 67
Const MA_basis_notes_const 				= 68
Const MSP_basis_of_elig_const 			= 69
Const MSP_basis_notes_const 			= 70
Const hc_eval_status					= 71
Const hc_eval_notes						= 72
Const pers_btn_one_const 				= 73
Const HC_major_prog_const				= 74
Const MSP_major_prog_const				= 75
Const LTC_waiver_notes_const			= 76
Const LTC_facility_notes_const			= 77
Const last_const						= 78

Dim HEALTH_CARE_MEMBERS()
ReDim HEALTH_CARE_MEMBERS(last_const, 0)

'Constants for the BILS_ARRAY
Const bils_ref_numb_const 		= 00
Const bils_date_const 			= 01
Const bils_service_code_const 	= 02
Const bils_service_info_const 	= 03
Const bils_gross_amt_const 		= 04
Const bils_third_payments_const = 05
Const bils_verif_code_const 	= 06
Const bils_verif_info_const 	= 07
Const bils_expense_type_code_const = 08
Const bils_expense_type_info_const = 09
Const bils_old_priority_const 	= 10
Const bils_depdnt_ind_const 	= 11
Const bils_hist_exp_type_code	= 12
Const bils_hist_exp_type_info	= 13
Const bils_hist_budg_period		= 14
Const bils_hist_budg_start		= 15
Const bils_hist_budg_end		= 16
Const bils_hist_monthly_used	= 17
Const bils_hist_monthly_unused	= 18
Const bils_hist_6_month_used	= 19
Const bils_hist_6_month_unused	= 20
Const bils_hist_sort_action		= 21
Const bils_hist_app_indc		= 22
Const bils_checkbox				= 23
Const bils_service_info_short_const = 24
Const last_bils_const 			= 25

Dim BILS_ARRAY()
ReDim BILS_ARRAY(last_bils_const, 0)

'defaulting some information
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
health_care_pending = False
health_care_active = False
hc_application_date = ""

form_selection_options = form_selection_options+chr(9)+"Health Care Programs Application for Certain Populations (DHS-3876)"
form_selection_options = form_selection_options+chr(9)+"MNsure Application for Health Coverage and Help paying Costs (DHS-6696)"
form_selection_options = form_selection_options+chr(9)+"Health Care Programs Renewal (DHS-3418)"
form_selection_options = form_selection_options+chr(9)+"Combined Annual Renewal for Certain Populations (DHS-3727)"
form_selection_options = form_selection_options+chr(9)+"Application for Payment of Long-Term Care Services (DHS-3531)"
form_selection_options = form_selection_options+chr(9)+"Renewal for People Receiving Medical Assistance for Long-Term Care Services (DHS-2128)"
form_selection_options = form_selection_options+chr(9)+"Breast and Cervical Cancer Coverage Group (DHS-3525)"
form_selection_options = form_selection_options+chr(9)+"Minnesota Family Planning Program Application (DHS-4740)"
form_selection_options = form_selection_options+chr(9)+"SAGE Enrollment Form (MA/BC PE Only)"
form_selection_options = form_selection_options+chr(9)+"Screen Our Circle Form (MA/BC PE Only)"
form_selection_options = form_selection_options+chr(9)+"Combined Six Month Report (DHS-5576)"
form_selection_options = form_selection_options+chr(9)+"No Form - Ex Parte Determination"

ma_basis_of_elig_list = "Disabled"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Blind"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Elderly"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Parent"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Caretaker"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Adult without Children"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Pregnant"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Child (19-20)"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Child (2-18)"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Child (0-1)"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Auto Newborn"
' ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Medical Emergency"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Foster Care Child"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Prev. Foster Care Child"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"Breast/Cervical Cancer"
ma_basis_of_elig_list = ma_basis_of_elig_list+chr(9)+"No Eligibility"

msp_basis_of_elig_list = "Disabled"
msp_basis_of_elig_list = msp_basis_of_elig_list+chr(9)+"Blind"
msp_basis_of_elig_list = msp_basis_of_elig_list+chr(9)+"Elderly"
msp_basis_of_elig_list = msp_basis_of_elig_list+chr(9)+"No Eligibility"
msp_basis_of_elig_list = msp_basis_of_elig_list+chr(9)+"No MEDICARE"

avs_form_status_list = "Select One..."
avs_form_status_list = avs_form_status_list+chr(9)+"Not Required"
avs_form_status_list = avs_form_status_list+chr(9)+"No Form Received"
avs_form_status_list = avs_form_status_list+chr(9)+"Incomplete - Received - Not Signed/Dated Correctly"
avs_form_status_list = avs_form_status_list+chr(9)+"Incomplete - Received - Not Signed by all Required Members"
avs_form_status_list = avs_form_status_list+chr(9)+"Complete - All Forms and Signatures Received"

page_display = ""
'These are the ways the pages of the dialog are selected, each is associated with a number
show_member_page 	= 0
show_jobs_page 		= 1
show_busi_page 		= 2
show_unea_page 		= 3
show_asset_page 	= 4
show_cars_rest_page	= 5
show_expenses_page 	= 6
show_other_page 	= 7
bils_page 			= 8
imig_page 			= 9
retro_page			= 10
verifs_page 		= 11
ltc_page			= 12
last_page 			= 13

'BUTTON definitions
'START AT 1000 OR ABOVE. Person buttons start at 500
Const hc_memb_btn 	= 1010
Const jobs_inc_btn 	= 1011
Const busi_inc_btn 	= 1012
Const unea_inc_btn 	= 1013
Const assets_btn 	= 1014
Const cars_rest_btn	= 1015
Const expenses_btn 	= 1016
Const other_btn 	= 1017
Const bils_btn		= 2018
Const imig_btn 		= 2019
Const verifs_page_btn 	= 2020
Const retro_btn		= 2021
Const ltc_page_btn	= 2022
Const last_btn		= 2030
Const verif_button	= 2500
Const clear_verifs_btn = 2510

Const completed_hc_eval_btn = 3000
Const next_btn				= 3010

Const instructions_btn = 5000
Const video_demo_btn = 5010

'We define a lot of things in dialogs, this makes sure they are available outside of the functions as well
Dim app_sig_status, app_sig_notes, client_delay_check, TIKL_check, MA_BC_end_of_cert_TIKL_check
Dim ma_bc_authorization_form, ma_bc_authorization_form_date, ma_bc_authorization_form_missing_checkbox
Dim bils_notes, verifs_needed, verif_req_form_sent_date, number_verifs_checkbox, case_details_notes
Dim last_page_numb
Dim retro_income_detail, retro_asset_detail, retro_expense_detail, ltc_elig_notes, ltc_info_still_needed
Dim retro_income_verifs_months, retro_asset_verifs_months, retro_expense_verifs_months, retro_notes
Dim avs_form_status, avs_form_notes, avs_portal_notes

'THE SCRIPT =====================================================================================================
EMConnect ""								'connect to BlueZone
Call check_for_MAXIS(False)					'Make sure we are in MAXIS
Call get_county_code						'Checking for the county
Call MAXIS_case_number_finder(MAXIS_case_number)		'Grabbing the CASE Number

If MAXIS_case_number <> "" Then				'If we know the CASE Number, we can attempt to read the form date
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
BeginDialog Dialog1, 0, 0, 400, 300, "Health Care Evaluation"
  EditBox 80, 200, 50, 15, MAXIS_case_number
  DropListBox 80, 220, 310, 45, "Select One..."+form_selection_options, HC_form_name
  DropListBox 300, 240, 90, 45, "No"+chr(9)+"Yes"+chr(9)+"N/A - Ex Parte Process", ltc_waiver_request_yn
  EditBox 80, 260, 50, 15, form_date
  EditBox 80, 280, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 285, 280, 50, 15
    CancelButton 340, 280, 50, 15
    PushButton 315, 35, 50, 13, "Instructions", instructions_btn
    PushButton 315, 50, 50, 13, "Video Demo", video_demo_btn
  Text 120, 10, 120, 10, "Health Care Evaluation Script"
  Text 40, 40, 255, 20, "This script is to be run once MAXIS STAT panels have been updated with all accurate information from a Health Care Application Form."
  Text 40, 65, 255, 25, "If information displayed in this script is inaccurate, this means the information entered into STAT requires update. Cancel the script run and update STAT panels before running the script again."
  Text 40, 95, 255, 10, "The information and coding in STAT will directly pull into the script details:"
  Text 55, 105, 250, 10, "- Panels coded as needing verification will show up as verifications needed."
  Text 55, 115, 250, 10, "- Income amounts will be pulled from JOBS / UNEA / BUSI / ect panels"
  Text 60, 125, 150, 10, "and cannot be updated in the script dialogs."
  Text 55, 135, 250, 10, "- Asset amounts will be pulled from ACCT / CASH / SECU / ect panels and"
  Text 60, 145, 175, 10, "cannot be updated in the script dialogs."
  Text 55, 155, 250, 10, "- The details in STAT Panels should be accurate and the script serves as a"
  Text 60, 165, 245, 10, "secondary review of information that makes an eligibility determinations."
  Text 35, 180, 300, 10, "IF THE CASE INFORMATION IS WRONG IN THE SCRIPT, IT IS WRONG IN THE SYSTEM"
  Text 25, 205, 50, 10, "Case Number:"
  Text 15, 225, 60, 10, "Health Care Form:"
  Text 115, 245, 185, 10, "Does this form qualify to request LTC/Waiver Services?"
  Text 25, 265, 40, 10, "Form Date:"
  Text 15, 285, 60, 10, "Worker Signature:"
  GroupBox 30, 25, 345, 170, "Health Care Processing"
  Text 135, 265, 110, 10, "For ex parte, use processing date"
EndDialog

DO
	DO
	   	err_msg = ""
	   	Dialog Dialog1
	   	cancel_without_confirmation

		If HC_form_name = "No Form - Ex Parte Determination" Then
			ltc_waiver_request_yn = "N/A - Ex Parte Process"
			form_date = date & ""
		End If
	    If ButtonPressed > 4000 Then
			If ButtonPressed = instructions_btn Then Call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20HEALTH%20CARE%20EVALUATION.docx")
			If ButtonPressed = video_demo_btn Then Call open_URL_in_browser("https://web.microsoftstream.com/video/21fa4c6c-0b95-4a53-b683-9b3bdce9fe95?referrer=https:%2F%2Fgbc-word-edit.officeapps.live.com%2F")
			err_msg = "LOOP"
		Else
			Call validate_MAXIS_case_number(err_msg, "*")
			If HC_form_name = "Select One..." Then err_msg = err_msg & vbCr & "* Select the form received that you are processing a Health Care evaluation from."
			If IsDate(form_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the form being processed was received."
			'Add validation to make sure LTC field is blank for ex parte
			If HC_form_name = "No Form - Ex Parte Determination" AND ltc_waiver_request_yn <> "N/A - Ex Parte Process" THEN err_msg = err_msg & vbCr & "* Select 'N/A - Ex Parte Process' for the LTC/Waiver Services field."
			If trim(worker_signature) = "" Then err_msg = err_msg & vbCr & "* Enter your name to sign your CASE/NOTE."

			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		End If
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Add ex parte SQL connections and script
If HC_form_name = "No Form - Ex Parte Determination" Then

	Const inc_panel_name 	= 0
	Const inc_ref_numb 		= 1
	Const inc_inst_numb 	= 2
	Const inc_type_code 	= 3
	Const inc_type_info 	= 4

	Const inc_verif 		= 20
	Const inc_start 		= 21
	Const inc_end 			= 22
	Const inc_update_date 	= 23
	Const inc_prosp_amt 	= 24
	Const inc_retro_amt 	= 25

	Const last_inc_const 	= 30

	Dim INCOME_ARRAY()
	ReDim INCOME_ARRAY(last_inc_const, 0)

	Const memb_ref_numb_const 	= 0
	Const memb_pmi_numb_const 	= 1
	Const memb_ssn_const 		= 2
	Const memb_ssn_dash_const	= 3
	Const memb_age_const 		= 4
	Const memb_name_const 		= 5
	Const memb_active_hc_const	= 6
	Const hc_prog_1				= 7
	Const hc_type_1				= 8
	Const hc_prog_2				= 9
	Const hc_type_2				= 10
	Const hc_prog_3				= 11
	Const hc_type_3				= 12
	Const memb_smi_numb_const	= 13

	Const MEDI_expt_exists_const= 20
	Const MEDI_update_date		= 21
	Const MEDI_Part_A_begin		= 22
	Const MEDI_Part_A_end		= 23
	Const MEDI_Part_B_begin		= 24
	Const MEDI_Part_B_end		= 25

	Const FACI_exists_const 	= 30
	Const Currently_in_FACI		= 31
	Const FACI_name				= 32
	Const FACI_date_in 			= 33
	Const FACI_date_out 		= 34
	Const FACI_type_code 		= 35
	Const FACI_type_info 		= 36
	Const FACI_vendor			= 37
	Const FACI_Waiver_type_code = 38
	Const FACI_Waiver_type_info = 39
	Const FACI_FS_ELIG_YN		= 40
	Const FACI_FS_Type_code		= 41
	Const FACI_FS_Type_info		= 42
	Const FACI_LTC_Inelig_reason_code = 43
	Const FACI_LTC_Inelig_reason_info = 44
	Const FACI_LTC_begin_date 	= 45
	Const FACI_cnty_approval_yn = 46
	Const FACI_approval_cnty	= 47

	Const DISA_expt_exists_const= 50
	Const DISA_begin_date 		= 51
	Const DISA_end_date 		= 52
	Const DISA_cert_begin_date	= 53
	Const DISA_cert_end_date	= 54
	Const DISA_HC_status_code 	= 55
	Const DISA_HC_status_info 	= 56
	Const DISA_waiver_code 		= 57
	Const DISA_waiver_info 		= 58

	Const PDED_exists_const 		= 70
	Const PDED_PICKLE_exists		= 71
	Const PDED_PICKLE_info			= 72
	Const PDED_PICKLE_detail		= 73
	Const PDED_PICKLE_thrshld_date	= 74
	Const PDED_PICKLE_curr_RSDI		= 75
	Const PDED_PICKLE_thrshld_RSDI	= 76
	Const PDED_PICKLE_dsrgd_amt		= 77
	Const PDED_DAC_exists 			= 78


	Const memb_last_const 		= 90

	Dim MEMBER_INFO_ARRAY()
	ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)

	MAXIS_footer_month = CM_plus_1_mo					'we are reading CM +1 for information for now.
	MAXIS_footer_year = CM_plus_1_yr
	SQL_Case_Number = right("00000000" & MAXIS_case_number, 8)

	Call back_to_SELF
	EMReadScreen MX_region, 10, 22, 48
	MX_region = trim(MX_region)
	If MX_region <> "TRAINING" Then
		'declare the SQL statement that will query the database
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [CaseNumber] = '" & SQL_Case_Number & "'"

		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		ex_parte_determination_from_SQL = objRecordSet("SelectExParte")
		ex_parte_phase_1_worker = objRecordSet("Phase1HSR")
		ex_parte_phase_2_worker = objRecordSet("Phase2HSR")
		If ex_parte_determination_from_SQL = 0 Then
			If IsNull(ex_parte_phase_2_worker) = True and ex_parte_phase_1_worker = user_ID_for_validation Then
				Change_ex_parte_determination_msg = MsgBox("It appears that you previously made a determination on Phase 1 for this case." & vbCr & vbCr & "Do you need to update the Ex Parte Determination?", vbQuestion + vbYesNo, "Update Ex Parte Determination")
				If Change_ex_parte_determination_msg = vbNo Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") was previously determined to not meet Ex Parte criteria.")
				' If Change_ex_parte_determination_msg = vbYes Then
			Else
				call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") was previously determined to not meet Ex Parte criteria.")
			End If
		End If
		review_month_from_SQL = objRecordSet("HCEligReviewDate")
		If review_month_from_SQL = "" Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") is not listed on the Ex Parte Data Table and cannot be processed as Ex Parte.")
		review_month_from_SQL = DateAdd("d", 0, review_month_from_SQL)
		Call convert_date_into_MAXIS_footer_month(review_month_from_SQL, er_month, er_year)
		If DateDiff("d", date, review_month_from_SQL) =< 0 Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") has a HC ER listed in the Ex Parte Data Table as " & er_month & "/" & er_year & ", which is in the past and cannot be processed as Ex Parte.")
		If DateDiff("m", date, review_month_from_SQL) > 3 Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") has a HC ER listed in the Ex Parte Data Table as " & er_month & "/" & er_year & ", which is too far in the future to be processed as Ex Parte.")
		ex_parte_phase = ""
		If DateDiff("m", date, review_month_from_SQL) = 1 Then
			ex_parte_phase = "Phase 2"
		Else
			ex_parte_phase = "Phase 1"
		End If
	Else

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 361, 55, "TRAINING Region Selections"
			DropListBox 125, 10, 225, 45, "Select One..,"+chr(9)+"Ex Parte Evaluation (Phase 1)"+chr(9)+"Ex Parte Approval (Phase 2)", phase_to_run
			EditBox 175, 30, 25, 15, er_month
			EditBox 205, 30, 25, 15, er_year
			ButtonGroup ButtonPressed
				OkButton 245, 30, 50, 15
				CancelButton 300, 30, 50, 15
			Text 5, 15, 115, 10, "What process do you want to run?"
			Text 5, 35, 160, 10, "What Ex Parte Renewal Month are you running?"
		EndDialog

		Do
			Do
				err_msg = ""
				dialog Dialog1
				cancel_confirmation

				If phase_to_run = "Select One..." Then err_msg = err_msg & vbCr & "* Select if you need to use Phase 1 or Phase 2 functionality."
				Call validate_footer_month_entry(er_month, er_year, err_msg, "*")

				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = ""
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		If phase_to_run = "Ex Parte Evaluation (Phase 1)" Then ex_parte_phase = "Phase 1"
		If phase_to_run = "Ex Parte Approval (Phase 2)" Then ex_parte_phase = "Phase 2"
		review_month_from_SQL = er_month & "/1/" & er_year
		review_month_from_SQL = DateAdd("d", 0, review_month_from_SQL)
	End If

	If ex_parte_phase = "Phase 1" Then
		phase_1_review_month = er_month & "/" & er_year
		correct_ex_parte_revw_month_code = right("00" & er_month, 2) & " 20" & right(er_year, 2)
		start_of_prep_month = DateAdd("m", -4, review_month_from_SQL)

		Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

		NO_HC_EXISTS = False
		EMReadScreen current_case_pw, 4, 21, 14
		If current_case_pw <> "X127" Then ex_parte_determination = "Case Transfered Out of County"
		If ma_case = False and msp_case = False and unknown_hc_pending = False Then
			NO_HC_EXISTS = True
		ElseIf is_this_priv = True Then

		Else

			Call navigate_to_MAXIS_screen("STAT", "SUMM")
			Call write_value_and_transmit("BGTX", 20, 71)
			Call MAXIS_background_check

			Call navigate_to_MAXIS_screen("ELIG", "HC  ")		'Navigate to ELIG/HC
			'Here we start at the top of ELIG/HC and read each row to find HC information
			hc_row = 8
			Do
				pers_type = ""		'blanking out variables so they don't carry over from loop to loop
				std = ""
				meth = ""
				waiv = ""

				'reading the main HC Elig information - member, program, status
				EMReadScreen read_ref_numb, 2, hc_row, 3
				EMReadScreen clt_hc_prog, 4, hc_row, 28
				EMReadScreen hc_prog_status, 6, hc_row, 50
				ref_row = hc_row
				Do while read_ref_numb = "  "				'this will read for the reference number if there are multiple programs for a single member
					ref_row = ref_row - 1
					EMReadScreen read_ref_numb, 2, ref_row, 3
				Loop

				If hc_prog_status = "ACTIVE" Then			'If HC is currently active, we need to read more details about the program/eligibility
					clt_hc_prog = trim(clt_hc_prog)			'formatting this to remove whitespace
					If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then		'these are non-hc persons

						Call write_value_and_transmit("X", hc_row, 26)									'opening the ELIG detail spans
						If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then		'If it is an MSP, we want to read the type only from a specific place
							elig_msp_prog = clt_hc_prog
							EMReadScreen pers_type, 2, 6, 56
						Else																			'These are MA type programs (not MSP)
							'Now we have to fund the current month in elig to get the current elig type
							col = 19
							Do
								EMReadScreen span_month, 2, 6, col										'reading the month in ELIG
								EMReadScreen span_year, 2, 6, col+3

								'if the span month matchest current month plus 1, we are going to grab elig from that month
								If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then
									EMReadScreen pers_type, 2, 12, col - 2								'reading the ELIG TYPE
									EMReadScreen std, 1, 12, col + 3
									EMReadScreen meth, 1, 13, col + 2
									EMReadScreen waiv, 1, 17, col + 2
									Exit Do																'leaving once we've found the information for this elig
								End If
								col = col + 11			'this goes to the next column
							Loop until col = 85			'This is off the page - if we hit this, we did NOT find the elig type in this elig result

							'If we hit 85, we did not get the information. So we are going to read it from the last budget month (most current)
							If col = 85 Then
								EMReadScreen pers_type, 2, 12, 72										'reading the ELIG TYPE
								EMReadScreen std, 1, 12, 77
								EMReadScreen meth, 1, 13, 76
								EMReadScreen waiv, 1, 17, 76
							End If
						End If
						PF3			'leaving the elig detail information

						'now we need to add the information we just read to the member array
						memb_known = False										'default that the member know is false
						For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)								'Looking at all the members known in the array
							If MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs) = read_ref_numb Then	'if the member reference from ELIG matches the ARRAY, we are going to add more elig details
								memb_known = True														'look we found a person
								If MEMBER_INFO_ARRAY(hc_prog_1, known_membs) = "" Then				'finding which area of the array is blank to save the elig information there
									MEMBER_INFO_ARRAY(hc_prog_1, known_membs) 		= clt_hc_prog
									MEMBER_INFO_ARRAY(hc_type_1, known_membs) 		= pers_type
								ElseIf MEMBER_INFO_ARRAY(hc_prog_2, known_membs) = "" Then
									MEMBER_INFO_ARRAY(hc_prog_2, known_membs) 		= clt_hc_prog
									MEMBER_INFO_ARRAY(hc_type_2, known_membs) 		= pers_type
								ElseIf MEMBER_INFO_ARRAY(hc_prog_3, known_membs) = "" Then
									MEMBER_INFO_ARRAY(hc_prog_3, known_membs) 		= clt_hc_prog
									MEMBER_INFO_ARRAY(hc_type_3, known_membs) 		= pers_type
								End If
							End If
						Next

						'If this is an unknown member, and has not been added to the array already, we need to add it
						If memb_known = False Then
							ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

							'setting personal information to the array
							MEMBER_INFO_ARRAY(memb_ref_numb_const, memb_count) = read_ref_numb
							MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
							MEMBER_INFO_ARRAY(hc_prog_1, memb_count) 		= trim(clt_hc_prog)
							MEMBER_INFO_ARRAY(hc_type_1, memb_count) 		= trim(pers_type)

							memb_count = memb_count + 1 	'incrementing the array counter up for the next loop
						End If

					End If
				End If
				hc_row = hc_row + 1												'now we go to the next row
				EMReadScreen next_ref_numb, 2, hc_row, 3						'read the next HC information to find when we've reeached the end of the list
				EMReadScreen next_maj_prog, 4, hc_row, 28
			Loop until next_ref_numb = "  " and next_maj_prog = "    "

			CALL back_to_SELF()													'going to STAT/MEMB - because there is misssing personal information for the members discovered in this way
			Do
				CALL navigate_to_MAXIS_screen("STAT", "MEMB")
				EMReadScreen memb_check, 4, 2, 48
			Loop until memb_check = "MEMB"

			at_least_one_hc_active = False										'this is a default to identify if HC is active on the case
			For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)					'loop through the member array
				Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 20, 76)		'navigate to the member for this instance of the array
				EMReadscreen last_name, 25, 6, 30								'read and cormat the name from MEMB
				EMReadscreen first_name, 12, 6, 63
				last_name = trim(replace(last_name, "_", "")) & " "
				first_name = trim(replace(first_name, "_", "")) & " "
				MEMBER_INFO_ARRAY(memb_name_const, known_membs) = first_name & " " & last_name
				MEMBER_INFO_ARRAY(memb_active_hc_const, known_membs) = False
				EMReadScreen PMI_numb, 8, 4, 46									'capturing the PMI number
				PMI_numb = trim(PMI_numb)
				MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) = right("00000000" & PMI_numb, 8)			'we have to format the pmi to match the data list format (8 digits with leading 0s included)
				EMReadScreen MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), 11, 7, 42							'catpturing the SSN
				MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), " ", "")
				MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), "_", "")
				If MEMBER_INFO_ARRAY(hc_prog_1, known_membs) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, known_membs) = True		'setting the variable that identifies there is HC active based on the ELIG read from HC/ELIG
				If MEMBER_INFO_ARRAY(hc_prog_2, known_membs) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, known_membs) = True
				If MEMBER_INFO_ARRAY(hc_prog_3, known_membs) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, known_membs) = True


			Next

			persons_list = " "
			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
			Call write_value_and_transmit("01", 20, 76)
			Do
				EMReadScreen cur_memb_ref_numb, 2, 4, 33
				use_memb_ref = ""
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)					'loop through the member array
					If MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs) = cur_memb_ref_numb Then use_memb_ref = known_membs
				Next
				If use_memb_ref = "" Then
					ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)
					MEMBER_INFO_ARRAY(memb_ref_numb_const, memb_count) = cur_memb_ref_numb
					use_memb_ref = memb_count
					memb_count = memb_count + 1
				End If

				EMReadscreen last_name, 25, 6, 30								'read and cormat the name from MEMB
				EMReadscreen first_name, 12, 6, 63
				last_name = trim(replace(last_name, "_", "")) & " "
				first_name = trim(replace(first_name, "_", "")) & " "
				MEMBER_INFO_ARRAY(memb_name_const, use_memb_ref) = first_name & " " & last_name
				MEMBER_INFO_ARRAY(memb_active_hc_const, use_memb_ref) = False
				EMReadScreen PMI_numb, 8, 4, 46									'capturing the PMI number
				PMI_numb = trim(PMI_numb)
				MEMBER_INFO_ARRAY(memb_pmi_numb_const, use_memb_ref) = right("00000000" & PMI_numb, 8)			'we have to format the pmi to match the data list format (8 digits with leading 0s included)
				EMReadScreen MEMBER_INFO_ARRAY(memb_ssn_const, use_memb_ref), 11, 7, 42							'catpturing the SSN
				MEMBER_INFO_ARRAY(memb_ssn_const, use_memb_ref) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, use_memb_ref), " ", "")
				MEMBER_INFO_ARRAY(memb_ssn_const, use_memb_ref) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, use_memb_ref), "_", "")
				If MEMBER_INFO_ARRAY(hc_prog_1, use_memb_ref) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, use_memb_ref) = True		'setting the variable that identifies there is HC active based on the ELIG read from HC/ELIG
				If MEMBER_INFO_ARRAY(hc_prog_2, use_memb_ref) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, use_memb_ref) = True
				If MEMBER_INFO_ARRAY(hc_prog_3, use_memb_ref) <> "" Then MEMBER_INFO_ARRAY(memb_active_hc_const, use_memb_ref) = True


				transmit
				EMReadScreen MEMB_end_check, 13, 24, 2
			LOOP Until MEMB_end_check = "ENTER A VALID"

			all_update_dates_are_current = True

			income_count = 0
			Call navigate_to_MAXIS_screen("STAT", "UNEA")
			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
				Call write_value_and_transmit("01", 20, 79)

				EMReadScreen version_numb, 1, 2, 78
				If version_numb <> "0" Then
					Do
						ReDim Preserve INCOME_ARRAY(last_inc_const, income_count)
						INCOME_ARRAY(inc_panel_name, income_count) = "UNEA"
						INCOME_ARRAY(inc_ref_numb, income_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)

						EMReadScreen INCOME_ARRAY(inc_inst_numb, 	income_count), 1, 2, 73
						EMReadScreen INCOME_ARRAY(inc_type_code, 	income_count), 2, 5, 37
						EMReadScreen INCOME_ARRAY(inc_type_info, 	income_count), 18, 5, 40
						EMReadScreen INCOME_ARRAY(inc_verif, 		income_count), 1, 5, 65
						EMReadScreen INCOME_ARRAY(inc_start, 		income_count), 8, 7, 37
						EMReadScreen INCOME_ARRAY(inc_end, 			income_count), 8, 7, 68
						EMReadScreen INCOME_ARRAY(inc_update_date, 	income_count), 8, 21, 55
						EMReadScreen INCOME_ARRAY(inc_retro_amt, 	income_count), 8, 18, 39
						EMReadScreen INCOME_ARRAY(inc_prosp_amt, 	income_count), 8, 18, 68

						INCOME_ARRAY(inc_inst_numb, income_count) = "0" & INCOME_ARRAY(inc_inst_numb, income_count)
						INCOME_ARRAY(inc_type_info, income_count) = trim(INCOME_ARRAY(inc_type_info, income_count))
						INCOME_ARRAY(inc_retro_amt, income_count) = trim(INCOME_ARRAY(inc_retro_amt, income_count))
						INCOME_ARRAY(inc_prosp_amt, income_count) = trim(INCOME_ARRAY(inc_prosp_amt, income_count))
						If INCOME_ARRAY(inc_prosp_amt, income_count) = "" Then INCOME_ARRAY(inc_prosp_amt, income_count) = "0.00"

						If INCOME_ARRAY(inc_start, income_count) = "__ __ __" Then INCOME_ARRAY(inc_start, income_count) = ""
						INCOME_ARRAY(inc_start, income_count) = replace(INCOME_ARRAY(inc_start, income_count), " ", "/")

						If INCOME_ARRAY(inc_end, income_count) = "__ __ __" Then INCOME_ARRAY(inc_end, income_count) = ""
						INCOME_ARRAY(inc_end, income_count) = replace(INCOME_ARRAY(inc_end, income_count), " ", "/")

						If INCOME_ARRAY(inc_update_date, income_count) = "" Then INCOME_ARRAY(inc_update_date, income_count) = ""
						INCOME_ARRAY(inc_update_date, income_count) = replace(INCOME_ARRAY(inc_update_date, income_count), " ", "/")

						If IsDate(INCOME_ARRAY(inc_update_date, income_count)) = False Then
							all_update_dates_are_current = False
						Else
							If DateDiff("d", INCOME_ARRAY(inc_update_date, income_count), start_of_prep_month) > 0 Then all_update_dates_are_current = False
						End If
						income_count = income_count + 1
						transmit
						Emreadscreen edit_check, 7, 24, 2
					LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
				End If
			Next

			Call navigate_to_MAXIS_screen("STAT", "JOBS")
			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
				Call write_value_and_transmit("01", 20, 79)

				EMReadScreen version_numb, 1, 2, 78
				If version_numb <> "0" Then
					Do
						ReDim Preserve INCOME_ARRAY(last_inc_const, income_count)
						INCOME_ARRAY(inc_panel_name, income_count) = "JOBS"
						INCOME_ARRAY(inc_ref_numb, income_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)

						EMReadScreen INCOME_ARRAY(inc_inst_numb, 	income_count), 1, 2, 73
						EMReadScreen INCOME_ARRAY(inc_type_code, 	income_count), 1, 5, 34
						EMReadScreen INCOME_ARRAY(inc_type_info, 	income_count), 13, 5, 36
						EMReadScreen INCOME_ARRAY(inc_verif, 		income_count), 1, 6, 34
						EMReadScreen INCOME_ARRAY(inc_start, 		income_count), 8, 9, 35
						EMReadScreen INCOME_ARRAY(inc_end, 			income_count), 8, 9, 49
						EMReadScreen INCOME_ARRAY(inc_update_date, 	income_count), 8, 21, 55
						EMReadScreen INCOME_ARRAY(inc_retro_amt, 	income_count), 8, 17, 38
						EMReadScreen INCOME_ARRAY(inc_prosp_amt, 	income_count), 8, 17, 67

						INCOME_ARRAY(inc_inst_numb, income_count) = "0" & INCOME_ARRAY(inc_inst_numb, income_count)
						INCOME_ARRAY(inc_type_info, income_count) = trim(INCOME_ARRAY(inc_type_info, 	income_count))
						INCOME_ARRAY(inc_retro_amt, income_count) = trim(INCOME_ARRAY(inc_retro_amt, income_count))
						INCOME_ARRAY(inc_prosp_amt, income_count) = trim(INCOME_ARRAY(inc_prosp_amt, income_count))
						If INCOME_ARRAY(inc_prosp_amt, income_count) = "" Then INCOME_ARRAY(inc_prosp_amt, income_count) = "0.00"

						If INCOME_ARRAY(inc_start, income_count) = "__ __ __" Then INCOME_ARRAY(inc_start, income_count) = ""
						INCOME_ARRAY(inc_start, income_count) = replace(INCOME_ARRAY(inc_start, income_count), " ", "/")

						If INCOME_ARRAY(inc_end, income_count) = "__ __ __" Then INCOME_ARRAY(inc_end, income_count) = ""
						INCOME_ARRAY(inc_end, income_count) = replace(INCOME_ARRAY(inc_end, income_count), " ", "/")

						If INCOME_ARRAY(inc_update_date, income_count) = "" Then INCOME_ARRAY(inc_update_date, income_count) = ""
						INCOME_ARRAY(inc_update_date, income_count) = replace(INCOME_ARRAY(inc_update_date, income_count), " ", "/")

						If IsDate(INCOME_ARRAY(inc_update_date, income_count)) = False Then
							all_update_dates_are_current = False
						Else
							If DateDiff("d", INCOME_ARRAY(inc_update_date, income_count), start_of_prep_month) > 0 Then all_update_dates_are_current = False
						End If
						income_count = income_count + 1
						transmit
						Emreadscreen edit_check, 7, 24, 2
					LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
				End If
			Next

			Call navigate_to_MAXIS_screen("STAT", "BUSI")
			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
				Call write_value_and_transmit("01", 20, 79)

				EMReadScreen version_numb, 1, 2, 78
				If version_numb <> "0" Then
					Do
						ReDim Preserve INCOME_ARRAY(last_inc_const, income_count)
						INCOME_ARRAY(inc_panel_name, income_count) = "BUSI"
						INCOME_ARRAY(inc_ref_numb, income_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)

						EMReadScreen INCOME_ARRAY(inc_inst_numb, 	income_count), 1, 2, 73
						EMReadScreen INCOME_ARRAY(inc_type_code, 	income_count), 2, 5, 37
						EMReadScreen INCOME_ARRAY(inc_start, 		income_count), 8, 5, 55
						EMReadScreen INCOME_ARRAY(inc_end, 			income_count), 8, 5, 72
						EMReadScreen INCOME_ARRAY(inc_update_date, 	income_count), 8, 21, 55
						EMReadScreen INCOME_ARRAY(inc_prosp_amt, 	income_count), 8, 12, 69

						Call write_value_and_transmit("X", 6, 26)
						EMReadScreen INCOME_ARRAY(inc_verif, income_count), 1, 13, 73
						PF3

						INCOME_ARRAY(inc_inst_numb, income_count) = "0" & INCOME_ARRAY(inc_inst_numb, income_count)
						INCOME_ARRAY(inc_prosp_amt, income_count) = trim(INCOME_ARRAY(inc_prosp_amt, income_count))
						If INCOME_ARRAY(inc_prosp_amt, income_count) = "" Then INCOME_ARRAY(inc_prosp_amt, income_count) = "0.00"

						If INCOME_ARRAY(inc_type_code, income_count) = "01" Then INCOME_ARRAY(inc_type_info, income_count) = "Farming"
						If INCOME_ARRAY(inc_type_code, income_count) = "02" Then INCOME_ARRAY(inc_type_info, income_count) = "Real Estate"
						If INCOME_ARRAY(inc_type_code, income_count) = "03" Then INCOME_ARRAY(inc_type_info, income_count) = "Home Product Sales"
						If INCOME_ARRAY(inc_type_code, income_count) = "04" Then INCOME_ARRAY(inc_type_info, income_count) = "Other Sales"
						If INCOME_ARRAY(inc_type_code, income_count) = "05" Then INCOME_ARRAY(inc_type_info, income_count) = "Personal Services"
						If INCOME_ARRAY(inc_type_code, income_count) = "06" Then INCOME_ARRAY(inc_type_info, income_count) = "Paper Route"
						If INCOME_ARRAY(inc_type_code, income_count) = "07" Then INCOME_ARRAY(inc_type_info, income_count) = "In Home Daycare"
						If INCOME_ARRAY(inc_type_code, income_count) = "08" Then INCOME_ARRAY(inc_type_info, income_count) = "Rental Income"
						If INCOME_ARRAY(inc_type_code, income_count) = "09" Then INCOME_ARRAY(inc_type_info, income_count) = "Other Self Employment"

						If INCOME_ARRAY(inc_start, income_count) = "__ __ __" Then INCOME_ARRAY(inc_start, income_count) = ""
						INCOME_ARRAY(inc_start, income_count) = replace(INCOME_ARRAY(inc_start, income_count), " ", "/")

						If INCOME_ARRAY(inc_end, income_count) = "__ __ __" Then INCOME_ARRAY(inc_end, income_count) = ""
						INCOME_ARRAY(inc_end, income_count) = replace(INCOME_ARRAY(inc_end, income_count), " ", "/")

						If INCOME_ARRAY(inc_update_date, income_count) = "" Then INCOME_ARRAY(inc_update_date, income_count) = ""
						INCOME_ARRAY(inc_update_date, income_count) = replace(INCOME_ARRAY(inc_update_date, income_count), " ", "/")

						If IsDate(INCOME_ARRAY(inc_update_date, income_count)) = False Then
							all_update_dates_are_current = False
						Else
							If DateDiff("d", INCOME_ARRAY(inc_update_date, income_count), start_of_prep_month) > 0 Then all_update_dates_are_current = False
						End If
						income_count = income_count + 1
						transmit
						Emreadscreen edit_check, 7, 24, 2
					LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
				End If
			Next

			'Navigate to MEDI panel for each reference number to determine if MEDI exists and Part A and Part B details
			Call navigate_to_MAXIS_screen("STAT", "MEDI")
			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb) = True Then
					Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76)
					EMReadScreen version_numb, 1, 2, 78
					If version_numb = "0" Then MEMBER_INFO_ARRAY(MEDI_expt_exists_const, each_memb) = False
					If version_numb <> "0" Then
						MEMBER_INFO_ARRAY(MEDI_expt_exists_const, each_memb) = True
						EMReadScreen MEMBER_INFO_ARRAY(MEDI_update_date, each_memb), 8, 21, 55
						If MEMBER_INFO_ARRAY(MEDI_update_date, each_memb) = "        " Then MEMBER_INFO_ARRAY(MEDI_update_date, each_memb) = ""
						MEMBER_INFO_ARRAY(MEDI_update_date, each_memb) = replace(MEMBER_INFO_ARRAY(MEDI_update_date, each_memb), " ", "/")

						Do
							PF20
							EMReadScreen end_of_list, 9, 24, 14
						Loop Until end_of_list = "LAST PAGE"
						row = 17
						Do
							EMReadScreen begin_dt_a, 8, row, 24 		'reads part a start date
							begin_dt_a = replace(begin_dt_a, " ", "/")	'reformatting with / for date
							If begin_dt_a = "__/__/__" Then begin_dt_a = "" 		'blank out if not a date

							EMReadScreen end_dt_a, 8, row, 35	'reads part a end date
							end_dt_a =replace(end_dt_a , " ", "/")		'reformatting with / for date
							If end_dt_a = "__/__/__" Then end_dt_a = ""					'blank out if not a date
							' MsgBox "end_dt_a - " & end_dt_a & vbCr & "begin_dt_a - " & begin_dt_a
							If end_dt_a <> "" or begin_dt_a <> "" Then
								MEMBER_INFO_ARRAY(MEDI_Part_A_begin, each_memb) = begin_dt_a
								MEMBER_INFO_ARRAY(MEDI_Part_A_end, each_memb) = end_dt_a
								Exit Do
							End If

							row = row - 1
							' MsgBox "PART A row - " & rowDosent_date_01
							If row = 14 Then
								PF19
								EMReadScreen begining_of_list, 10, 24, 14
								' MsgBox "begining_of_list - " & begining_of_list & vbcr & "1"
								If begining_of_list = "FIRST PAGE" Then
									Exit Do
								Else
									row = 17
								End If
							End If
						Loop
						Do
							PF19
							EMReadScreen begining_of_list, 10, 24, 14
							' MsgBox "begining_of_list - " & begining_of_list & vbcr & "2"
						Loop Until begining_of_list = "FIRST PAGE"

						Do
							PF20
							EMReadScreen end_of_list, 9, 24, 14
							' MsgBox end_of_list & " - 2"
						Loop Until end_of_list = "LAST PAGE"
						row = 17
						Do
							EMReadScreen begin_dt_b, 8, row, 54 		'reads part a start date
							begin_dt_b = replace(begin_dt_b, " ", "/")	'reformatting with / for date
							If begin_dt_b = "__/__/__" Then begin_dt_b = "" 		'blank out if not a date

							EMReadScreen end_dt_b, 8, row, 65	'reads part a end date
							end_dt_b =replace(end_dt_b , " ", "/")		'reformatting with / for date
							If end_dt_b = "__/__/__" Then end_dt_b = ""					'blank out if not a date

							' MsgBox "end_dt_b - " & end_dt_b & vbCr & "begin_dt_b - " & begin_dt_b
							If end_dt_b <> "" or begin_dt_b <> "" Then
								MEMBER_INFO_ARRAY(MEDI_Part_B_begin, each_memb) = begin_dt_b
								MEMBER_INFO_ARRAY(MEDI_Part_B_end, each_memb) = end_dend_dt_bt_a
								Exit Do
							End If

							row = row - 1
							' MsgBox "PART B row - " & row
							If row = 14 Then
								PF19
								EMReadScreen begining_of_list, 10, 24, 14
								' MsgBox "begining_of_list - " & begining_of_list & vbcr & "3"
								If begining_of_list = "FIRST PAGE" Then
									Exit Do
								Else
									row = 17
								End If
							End If
						Loop
					End If
				End If
			Next


			call navigate_to_MAXIS_screen("STAT", "FACI")

			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb) = True Then
					Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76)
					EMReadScreen version_numb, 1, 2, 78
					If version_numb = "0" Then MEMBER_INFO_ARRAY(FACI_exists_const, each_memb) = False
					If version_numb <> "0" Then
						MEMBER_INFO_ARRAY(FACI_exists_const, each_memb) = True
						Do
							EMReadScreen FACI_current_panel, 1, 2, 73
							EMReadScreen FACI_total_check, 1, 2, 78
							EMReadScreen in_year_check_01, 4, 14, 53
							EMReadScreen in_year_check_02, 4, 15, 53
							EMReadScreen in_year_check_03, 4, 16, 53
							EMReadScreen in_year_check_04, 4, 17, 53
							EMReadScreen in_year_check_05, 4, 18, 53
							EMReadScreen out_year_check_01, 4, 14, 77
							EMReadScreen out_year_check_02, 4, 15, 77
							EMReadScreen out_year_check_03, 4, 16, 77
							EMReadScreen out_year_check_04, 4, 17, 77
							EMReadScreen out_year_check_05, 4, 18, 77
							If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or _
							(in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
								MEMBER_INFO_ARRAY(Currently_in_FACI, each_memb) = True
								If in_year_check_01 <> "____" and out_year_check_01 = "____" Then faci_row = 14
								If in_year_check_02 <> "____" and out_year_check_02 = "____" Then faci_row = 15
								If in_year_check_03 <> "____" and out_year_check_03 = "____" Then faci_row = 16
								If in_year_check_04 <> "____" and out_year_check_04 = "____" Then faci_row = 17
								If in_year_check_05 <> "____" and out_year_check_05 = "____" Then faci_row = 18

								EMReadScreen MEMBER_INFO_ARRAY(FACI_date_in, each_memb), 10, faci_row, 47
								EMReadScreen MEMBER_INFO_ARRAY(FACI_date_out, each_memb), 10, faci_row, 	71

								If MEMBER_INFO_ARRAY(FACI_date_in, each_memb) = "__ __ ____" Then MEMBER_INFO_ARRAY(FACI_date_in, each_memb) = ""
								MEMBER_INFO_ARRAY(FACI_date_in, each_memb) = replace(MEMBER_INFO_ARRAY(FACI_date_in, each_memb), " ", "/")
								If MEMBER_INFO_ARRAY(FACI_date_out, each_memb) = "__ __ ____" Then MEMBER_INFO_ARRAY(FACI_date_out, each_memb) = ""
								MEMBER_INFO_ARRAY(FACI_date_out, each_memb) = replace(MEMBER_INFO_ARRAY(FACI_date_out, each_memb), " ", "/")

								exit do
							Elseif FACI_current_panel = FACI_total_check then
								MEMBER_INFO_ARRAY(Currently_in_FACI, each_memb) = False
								exit do
							Else
								transmit
							End if
						Loop until FACI_current_panel = FACI_total_check

						If MEMBER_INFO_ARRAY(Currently_in_FACI, each_memb) = True then
							EMReadScreen MEMBER_INFO_ARRAY(FACI_name, each_memb), 30, 6, 43
							EMReadScreen MEMBER_INFO_ARRAY(FACI_type_code, each_memb), 2, 7, 43
							EmReadscreen MEMBER_INFO_ARRAY(FACI_vendor, each_memb), 8, 5, 43
							'List of FACI types
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "41" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "NF-I"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "42" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "NF-II"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "43" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "ICF-DD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "44" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Short stay in NF-I"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "45" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Short stay in NF-II"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "46" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Short stay in ICF-DD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "47" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "RTC - Not IMD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "48" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Medical Hospital"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "49" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "MSOP"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "50" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "IMD/RTC"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "51" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Rule 31 CD_IMD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "52" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Rule 36 MI-IMD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "53" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "IMD Hospitals"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "55" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Adult Foster Care/Rule 203"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "56" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "GRH (Not FC or Rule 36)"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "57" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Rule 36 MI - Non-IMD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "60" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Non-GRH"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "61" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Rule 31 CD - Non-IMD"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "67" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Family Violence Shelter"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "68" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "County Correctional Facility"
							IF MEMBER_INFO_ARRAY(FACI_type_code, each_memb) = "69" then MEMBER_INFO_ARRAY(FACI_type_info, each_memb) = "Non-Cty Adult Correctional"

							MEMBER_INFO_ARRAY(FACI_name, each_memb) = trim(replace(MEMBER_INFO_ARRAY(FACI_name, each_memb), "_", ""))
							MEMBER_INFO_ARRAY(FACI_vendor, each_memb) = trim(replace(MEMBER_INFO_ARRAY(FACI_vendor, each_memb), "_", ""))

							EMReadScreen MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb), 2, 7, 71
							EMReadScreen MEMBER_INFO_ARRAY(FACI_FS_ELIG_YN, each_memb), 1, 8, 43
							EMReadScreen MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb), 1, 8, 71
							EMReadScreen MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_code, each_memb), 1, 9, 43
							EMReadScreen MEMBER_INFO_ARRAY(FACI_LTC_begin_date, each_memb), 10, 10, 52
							EMReadScreen MEMBER_INFO_ARRAY(FACI_cnty_approval_yn, each_memb), 1, 12, 52
							EMReadScreen MEMBER_INFO_ARRAY(FACI_approval_cnty, each_memb), 2, 12, 71

							If MEMBER_INFO_ARRAY(FACI_LTC_begin_date, each_memb) = "__ __ ____" Then MEMBER_INFO_ARRAY(FACI_LTC_begin_date, each_memb) = ""
							MEMBER_INFO_ARRAY(FACI_LTC_begin_date, each_memb) = replace(MEMBER_INFO_ARRAY(FACI_LTC_begin_date, each_memb), " ", "/")

							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "__" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = ""
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "01" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "CADI"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "02" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "CAC"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "03" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "EW Single"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "04" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "EW Married"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "05" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "TBI"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "06" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "DD"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "07" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "ACS"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "08" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "SISEW Single"
							If MEMBER_INFO_ARRAY(FACI_Waiver_type_code, each_memb) = "09" Then MEMBER_INFO_ARRAY(FACI_Waiver_type_info, each_memb) = "SISEW Married"

							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "_" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = ""
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "1" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Federally Subsidized Housing for Elderly"
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "2" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Licensed Facility/Treatment Center for Chemical Dependency"
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "3" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Blind or Disabled RSDI/SSI Recipient"
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "4" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Family Violence Shelter"
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "5" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Temporary Shelter for Homeless"
							If MEMBER_INFO_ARRAY(FACI_FS_Type_code, each_memb) = "6" Then  MEMBER_INFO_ARRAY(FACI_FS_Type_info, each_memb) = "Not a facility by FS Definition"

							If MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_code, each_memb) = "_" Then MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_info, each_memb) = ""
							If MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_code, each_memb) = "L" Then MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_info, each_memb) = "This Level of Care Not Required"
							If MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_code, each_memb) = "N" Then MEMBER_INFO_ARRAY(FACI_LTC_Inelig_reason_info, each_memb) = "Not pre-Screened"

						End if
					End If
				End If
			Next

			CALL navigate_to_MAXIS_screen("STAT", "DISA")

			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "No HC DISA Coded"
				MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "No Waiver Coded"
				If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb) = True Then
					Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76)
					EMReadScreen version_numb, 1, 2, 78
					If version_numb = "0" Then MEMBER_INFO_ARRAY(DISA_expt_exists_const, each_memb) = False
					If version_numb <> "0" Then
						MEMBER_INFO_ARRAY(DISA_expt_exists_const, each_memb) = True

						EMReadScreen MEMBER_INFO_ARRAY(DISA_begin_date, each_memb), 10, 6, 47
						MEMBER_INFO_ARRAY(DISA_begin_date, each_memb) = replace(MEMBER_INFO_ARRAY(DISA_begin_date, each_memb), " ", "/")
						If MEMBER_INFO_ARRAY(DISA_begin_date, each_memb) = "__/__/____" Then MEMBER_INFO_ARRAY(DISA_begin_date, each_memb) = ""

						EMReadScreen MEMBER_INFO_ARRAY(DISA_end_date, each_memb), 10, 6, 69
						MEMBER_INFO_ARRAY(DISA_end_date, each_memb) = replace(MEMBER_INFO_ARRAY(DISA_end_date, each_memb), " ", "/")
						If MEMBER_INFO_ARRAY(DISA_end_date, each_memb) = "__/__/____" Then MEMBER_INFO_ARRAY(DISA_end_date, each_memb) = ""

						EMReadScreen MEMBER_INFO_ARRAY(DISA_cert_begin_date, each_memb), 10, 7, 47
						MEMBER_INFO_ARRAY(DISA_cert_begin_date, each_memb) = replace(MEMBER_INFO_ARRAY(DISA_cert_begin_date, each_memb), " ", "/")
						If MEMBER_INFO_ARRAY(DISA_cert_begin_date, each_memb) = "__/__/____" Then MEMBER_INFO_ARRAY(DISA_cert_begin_date, each_memb) = ""

						EMReadScreen MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb), 10, 7, 69
						MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb) = replace(MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb), " ", "/")
						If MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb) = "__/__/____" Then MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb) = ""

						EMReadScreen MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) , 2, 13, 59
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "01" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "RSDI Only Disability"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "02" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "RSDI Only Blindness"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "03" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "SSI, SSI/RSDI Disability"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "04" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "SSI, SSI/RSDI Blindness"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "06" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "SMRT Pend or SSA Pend"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "08" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "Certified Blind"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "10" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "Certified Disabled"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "11" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "Special Category - Disabled Child"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "20" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "TEFRA - Disabled"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "21" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "TEFRA - Blind"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "22" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "MA-EPD"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "23" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "MA/Waiver"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "24" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "SSA/SMRT Appeal Pending"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "26" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "SSA/SMRT Disa Deny"
						If MEMBER_INFO_ARRAY(DISA_HC_status_code, each_memb) = "__" Then MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb) = "No HC DISA Status"

						EMReadScreen MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb), 1, 14, 59
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "F" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC CADI Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "G" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC CADI DIversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "H" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC CAC Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "I" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC CAC Diversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "J" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC EW Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "K" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC EW Diversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "L" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC TBI NF Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "M" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC TBI NF Diversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "P" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC TBI NB Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "Q" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "LTC TBI NB Diversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "R" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "DD Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "S" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "DD Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "Y" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "CSG Conversion"
						If MEMBER_INFO_ARRAY(DISA_waiver_code, each_memb) = "_" Then MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb) = "No Waiver"

					End If
				End If
			Next

			call navigate_to_MAXIS_screen("STAT", "PDED")

			For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb) = True Then
					Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76)
					EMReadScreen version_numb, 1, 2, 78
					If version_numb = "0" Then MEMBER_INFO_ARRAY(PDED_exists_const, each_memb) = False
					If version_numb <> "0" Then
						MEMBER_INFO_ARRAY(PDED_exists_const, each_memb) = True
						MEMBER_INFO_ARRAY(PDED_PICKLE_exists, each_memb) = False
						MEMBER_INFO_ARRAY(PDED_DAC_exists , each_memb) = False

						EMReadScreen pickle_dsrgd_yn, 1, 6, 60
						If pickle_dsrgd_yn = "Y" Then
							MEMBER_INFO_ARRAY(PDED_PICKLE_exists, each_memb) = True
							Call write_value_and_transmit("X", 6, 40)
							EMReadScreen MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_date, each_memb), 8, 5, 48
							EMReadScreen MEMBER_INFO_ARRAY(PDED_PICKLE_curr_RSDI, each_memb), 8, 6, 48
							EMReadScreen MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_RSDI, each_memb), 8, 7, 48
							EMReadScreen MEMBER_INFO_ARRAY(PDED_PICKLE_dsrgd_amt, each_memb), 8, 8, 48

							MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_date, each_memb) = replace(MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_date, each_memb), " ", "/")
							MEMBER_INFO_ARRAY(PDED_PICKLE_curr_RSDI, each_memb) = trim(MEMBER_INFO_ARRAY(PDED_PICKLE_curr_RSDI, each_memb))
							MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_RSDI, each_memb) = trim(MEMBER_INFO_ARRAY(PDED_PICKLE_thrshld_RSDI, each_memb))
							MEMBER_INFO_ARRAY(PDED_PICKLE_dsrgd_amt, each_memb) = trim(MEMBER_INFO_ARRAY(PDED_PICKLE_dsrgd_amt, each_memb))

							PF3
						End If

						EMReadScreen dac_dsrgd_yn, 1, 8, 60
						If dac_dsrgd_yn = "Y" Then MEMBER_INFO_ARRAY(PDED_DAC_exists , each_memb) = True
					End If
				End If
			Next


		End If

		Call navigate_to_MAXIS_screen("CASE", "CURR")

		instructions_button = 1000
		policy_1_button = 1010
		policy_2_button = 1020
		policy_3_button = 1030

		If NO_HC_EXISTS = True Then ex_parte_determination = "Health Care has been Closed"

		Dialog1 = ""

		BeginDialog Dialog1, 0, 0, 556, 385, "Phase 1 - Ex Parte Evaluation"
			GroupBox 10, 280, 505, 85, "Ex Parte Evaluation"
			Text 20, 295, 80, 10, "Ex Parte Evaluation:"
			DropListBox 100, 290, 130, 45, ""+chr(9)+"Appears Ex Parte"+chr(9)+"Cannot be Processed as Ex Parte"+chr(9)+"Health Care has been Closed"+chr(9)+"Case Transfered Out of County", ex_parte_determination
			Text 235, 295, 200, 10, "Identifying a case as NOT Ex Parte requires explanation."
			Text 15, 315, 90, 10, "Not Ex Parte Explanation:"
			ComboBox 100, 310, 400, 45, "Select or Enter Reason for NOT Ex Parte"+chr(9)+"No spenddown currently approved and Income indicates a spenddown may be required."+chr(9)+"Income cannot be verified without resident interaction."+chr(9)+"Not all household members on HC meed an ABD basis."+chr(9)+"Resident is not in compliance with SSA."+chr(9)+"Resident is not in compliance with OMB/PBEN.", ex_parte_denial_select
			Text 20, 335, 80, 10, "Not Ex Parte Notes:"
			EditBox 100, 330, 400, 15, ex_parte_denial_notes 'ex_parte_denial_explanation
			Text 15, 350, 495, 10, "If 'Not Ex Parte' you must enter Explanation and/or Notes. You do not have to enter both. The total character limit is 255 for the combination of all information."
			Text 15, 365, 70, 10, "Worker Signature:"
			EditBox 80, 360, 110, 15, worker_signature
			' GroupBox 10, 0, 455, 25, "Case Information"
			If current_case_pw <> "X127" Then Text 10, 5, 150, 10, "CASE IS NOT IN HENNEPIN COUNTY - (" & right(current_case_pw, 2) & ")"
			Text 275, 5, 75, 10, "Case Number: " & MAXIS_case_number
			' Text 65, 10, 70, 10, MAXIS_case_number
			Text 375, 5, 75, 10, "Review Month: " & phase_1_review_month
			Text 280, 15, 125, 10, "SNAP Status: " & snap_status
			Text 283, 25, 125, 10, "MFIP Status: " & mfip_status
			ButtonGroup ButtonPressed
				OkButton 440, 365, 50, 15
				CancelButton 500, 365, 50, 15
				Text 480, 5, 70, 10, "--- INSTRUCTIONS ---"
				PushButton 475, 15, 80, 15, "Instructions", instructions_button
				Text 495, 40, 45, 10, "--- POLICY ---"
				PushButton 475, 50, 80, 15, "DHS #23-21-18", policy_1_button
				' PushButton 490, 65, 55, 15, policy_2, policy_2_button
				' PushButton 490, 80, 55, 15, policy_3, policy_3_button
				Text 490, 105, 55, 10, "--- NAVIGATE ---"
				PushButton 485, 117, 25, 10, "ACCI", acci_button
				PushButton 515, 117, 25, 10, "BILS", bils_button
				PushButton 485, 130, 25, 10, "BUDG", budg_button
				PushButton 515, 130, 25, 10, "BUSI", busi_button
				PushButton 485, 143, 25, 10, "DISA", disa_button
				PushButton 515, 143, 25, 10, "EMMA", emma_button
				PushButton 485, 156, 25, 10, "FACI", faci_button
				PushButton 515, 156, 25, 10, "HCMI", hcmi_button
				PushButton 485, 169, 25, 10, "IMIG", imig_button
				PushButton 515, 169, 25, 10, "INSA", insa_button
				PushButton 485, 182, 25, 10, "JOBS", jobs_button
				PushButton 515, 182, 25, 10, "LUMP", lump_button
				PushButton 485, 195, 25, 10, "MEDI", medi_button
				PushButton 515, 195, 25, 10, "MEMB", memb_button
				PushButton 485, 208, 25, 10, "MEMI", memi_button
				PushButton 515, 208, 25, 10, "PBEN", pben_button
				PushButton 485, 221, 25, 10, "PDED", pded_button
				PushButton 515, 221, 25, 10, "REVW", revw_button
				PushButton 485, 234, 25, 10, "SPON", spon_button
				PushButton 515, 234, 25, 10, "STWK", stwk_button
				PushButton 485, 247, 25, 10, "UNEA", unea_button

			If NO_HC_EXISTS = True Then
				Text 15, 25, 200, 10, "No Health Care Programs Active or Pending."
			Else
				y_pos = 10
				If all_update_dates_are_current = False Then
					Text 15, y_pos, 255, 10, "* * * * INCOME HAS NOT BEEN UPDATED DURRING PREP MONTH * * * *"
					Text 55, y_pos+10, 185, 10, "* * * * Manual Review of this case is important. * * * *"
					y_pos = y_pos + 25
				Else
					y_pos = 15
				End If
				x_pos = 10
				memb_with_hc_list = " "
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					If y_pos > 35 Then y_pos = 35
					If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb) = True Then
						memb_with_hc_list = memb_with_hc_list & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " "
						memb_grp_x = x_pos
						memb_grp_y = y_pos
						x_pos = x_pos + 10
						y_pos = y_pos +10

						Text x_pos+100, y_pos, 65, 10, "PMI: " & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						Text x_pos, y_pos, 150, 10, "Health Care Programs:"
						y_pos = y_pos + 10
						If MEMBER_INFO_ARRAY(hc_prog_1, each_memb) <> "" Then
							Text x_pos+10, y_pos, 65, 10, MEMBER_INFO_ARRAY(hc_prog_1, each_memb) & " - " & MEMBER_INFO_ARRAY(hc_type_1, each_memb)
							y_pos = y_pos + 10
						End If
						If MEMBER_INFO_ARRAY(hc_prog_2, each_memb) <> "" Then
							Text x_pos+10, y_pos, 65, 10, MEMBER_INFO_ARRAY(hc_prog_2, each_memb) & " - " & MEMBER_INFO_ARRAY(hc_type_2, each_memb)
							y_pos = y_pos + 10
						End If
						If MEMBER_INFO_ARRAY(hc_prog_3, each_memb) <> "" Then
							Text x_pos+10, y_pos, 65, 10, MEMBER_INFO_ARRAY(hc_prog_3, each_memb) & " - " & MEMBER_INFO_ARRAY(hc_type_3, each_memb)
							y_pos = y_pos + 10
						End If
						y_pos = y_pos + 5

						Text x_pos, y_pos, 200, 10, "DISA:        " & MEMBER_INFO_ARRAY(DISA_HC_status_info, each_memb)
						y_pos = y_pos + 10
						If MEMBER_INFO_ARRAY(DISA_end_date, each_memb) <> "" Then
							Text x_pos+35, y_pos, 100, 10, "End Date: " & MEMBER_INFO_ARRAY(DISA_end_date, each_memb)
							y_pos = y_pos + 10
						ElseIf MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb) <> "" Then
							Text x_pos+35, y_pos, 100, 10, "CERT End Date: " & MEMBER_INFO_ARRAY(DISA_cert_end_date, each_memb)
							y_pos = y_pos + 10
						End If
						y_pos = y_pos + 5
						Text x_pos, y_pos, 100, 10, "Waiver:     " & MEMBER_INFO_ARRAY(DISA_waiver_info, each_memb)
						y_pos = y_pos + 15
						If MEMBER_INFO_ARRAY(MEDI_expt_exists_const, each_memb) = True Then
							MEDI_part_a = ""
							If MEMBER_INFO_ARRAY(MEDI_Part_A_end, each_memb) <> "" or MEMBER_INFO_ARRAY(MEDI_Part_A_begin, each_memb) <> "" Then
								If MEMBER_INFO_ARRAY(MEDI_Part_A_begin, each_memb) <> "" Then MEDI_part_a = "Start: " & MEMBER_INFO_ARRAY(MEDI_Part_A_begin, each_memb)
								If MEMBER_INFO_ARRAY(MEDI_Part_A_end, each_memb) <> "" Then MEDI_part_a = MEDI_part_b & ", End: " & MEMBER_INFO_ARRAY(MEDI_Part_A_end, each_memb)
								If left(MEDI_part_a, 1) = "," Then MEDI_part_a = right(MEDI_part_a, len(MEDI_part_a)-2)
							End If
							MEDI_part_b = ""
							If MEMBER_INFO_ARRAY(MEDI_Part_B_end, each_memb) <> "" or MEMBER_INFO_ARRAY(MEDI_Part_B_begin, each_memb) <> "" Then
								If MEMBER_INFO_ARRAY(MEDI_Part_B_begin, each_memb) <> "" Then MEDI_part_b = "Start: " & MEMBER_INFO_ARRAY(MEDI_Part_B_begin, each_memb)
								If MEMBER_INFO_ARRAY(MEDI_Part_B_end, each_memb) <> "" Then MEDI_part_b = MEDI_part_b & ", End: " & MEMBER_INFO_ARRAY(MEDI_Part_B_end, each_memb)
								If left(MEDI_part_b, 1) = "," Then MEDI_part_b = right(MEDI_part_b, len(MEDI_part_b)-2)
							End If
							Text x_pos, y_pos, 200, 10, "MEDI Info: Part A: " & MEDI_part_a
							Text x_pos+35, y_pos+10, 200, 10, " Part B: " & MEDI_part_b
							y_pos = y_pos + 25
						Else
							Text x_pos, y_pos, 200, 10, "NO MEDI Panel entered for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
							y_pos = y_pos + 10
						End If

						If MEMBER_INFO_ARRAY(Currently_in_FACI, each_memb) = True then
							Text x_pos, y_pos, 200, 10, MEMBER_INFO_ARRAY(memb_name_const, each_memb) & " - currently in a facility"
							Text x_pos+10, y_pos+10, 200, 10, "Facility Name: " & MEMBER_INFO_ARRAY(FACI_name, each_memb)
							Text x_pos+20, y_pos+20, 150, 10, "Entry Date: " & MEMBER_INFO_ARRAY(FACI_date_in, each_memb)
							Text x_pos+14, y_pos+30, 200, 10,"Facility Type: " & MEMBER_INFO_ARRAY(FACI_type_info, each_memb)
							y_pos = y_pos + 45
						Else
							Text x_pos, y_pos, 200, 10, MEMBER_INFO_ARRAY(memb_name_const, each_memb) & " - Does not appear to be in a facility"
							y_pos = y_pos + 10
						End If

						If MEMBER_INFO_ARRAY(PDED_PICKLE_exists, each_memb) = True Then
							Text x_pos, y_pos, 200, 10, "PICKLE Disregard of $ " & MEMBER_INFO_ARRAY(PDED_PICKLE_dsrgd_amt, each_memb)
							y_pos = y_pos + 10
						End If
						If MEMBER_INFO_ARRAY(PDED_DAC_exists , each_memb) = True Then
							Text x_pos, y_pos, 200, 10, "DAC Disregard Exists"
							y_pos = y_pos + 10
						End If

						' If all_update_dates_are_current = True Then
							first_income = True
							For each_income = 0 to UBound(INCOME_ARRAY, 2)
								If INCOME_ARRAY(inc_ref_numb, each_income) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) Then
									If first_income = True Then
										y_pos = y_pos + 10
										Text x_pos, y_pos, 200, 10, "INCOME for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
										y_pos = y_pos + 10
										first_income = False
									End If
									Text x_pos+5, y_pos, 150, 10, INCOME_ARRAY(inc_panel_name, each_income) & " Income - " & INCOME_ARRAY(inc_type_info, each_income)
									Text x_pos+160, y_pos, 50, 10, "Verif: " & INCOME_ARRAY(inc_verif, each_income)
									Text x_pos+10, y_pos+10, 100, 10, "Gross Income $ " & INCOME_ARRAY(inc_prosp_amt, each_income)
									y_pos = y_pos + 20
									If INCOME_ARRAY(inc_end, each_income) <> "" Then
										Text x_pos+10, y_pos, 200, 10, "Income has ended as of " & INCOME_ARRAY(inc_end, each_income)
										y_pos = y_pos + 10
									End If
								End If
							Next
						' End If
						If last_y_pos = "" Then
							last_y_pos = y_pos
						Else
							If y_pos > last_y_pos Then last_y_pos = y_pos
						End If
						GroupBox memb_grp_x, memb_grp_y, 220, y_pos-memb_grp_y+5, "MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						x_pos = x_pos + 230
					End If
				Next

				first_income = True
				y_pos = last_y_pos + 20
				For each_income = 0 to UBound(INCOME_ARRAY, 2)
					If InStr(memb_with_hc_list, INCOME_ARRAY(inc_ref_numb, each_income)) = 0 Then
						If first_income = True Then
							grp_box_start = y_pos
							y_pos = y_pos + 15
							first_income = False
						End If
						Text 20, y_pos, 200, 10, "MEMB " & INCOME_ARRAY(inc_ref_numb, each_income) & " - " & INCOME_ARRAY(inc_panel_name, each_income) & " Income - Gross Amount $ " & INCOME_ARRAY(inc_prosp_amt, each_income) & " - Verif: " & INCOME_ARRAY(inc_verif, each_income)
						y_pos = y_pos + 10
						If INCOME_ARRAY(inc_end, each_income) <> "" Then
							Text 30, y_pos, 200, 10, "Income has ended as of " & INCOME_ARRAY(inc_end, each_income)
							y_pos = y_pos + 10
						End If
					End If
				Next
				If first_income = False Then GroupBox 10, grp_box_start, 450, y_pos-grp_box_start+5, "INCOME for Houeshold Members not on HC on this case"

			End If
		EndDialog

		'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
		DO
			Do
				err_msg = ""    'This is the error message handling
				Dialog Dialog1
				cancel_confirmation
				'Function belows creates navigation to STAT panels for navigation buttons
				MAXIS_dialog_navigation

				ex_parte_denial_explanation = ""
				ex_parte_denial_explanation = trim(replace(ex_parte_denial_select, "Select or Enter Reason for NOT Ex Parte", ""))
				ex_parte_denial_explanation = ex_parte_denial_explanation & " " & trim(ex_parte_denial_notes)
				ex_parte_denial_explanation = trim(ex_parte_denial_explanation)
				' MsgBox ex_parte_denial_explanation & vbCr & vbCr & len(ex_parte_denial_explanation)


				'Add placeholder link to script instructions - To DO - update with correct link
				If ButtonPressed = instructions_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20HEALTH%20CARE%20EVALUATION%20-%20EX%20PARTE%20PROCESS.docx"

				'Add placeholder links for policy buttons - TO DO - update with correct links
				If ButtonPressed = policy_1_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=mndhs-062948"
				' If ButtonPressed = policy_2_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
				' If ButtonPressed = policy_3_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"


				'Add validation to ensure ex parte determination is made
				If ex_parte_determination = "" THEN err_msg = err_msg & vbCr & "* You must make an ex parte determination."

				'Add validation that if ex parte approved, then explanation should be blank
				If ex_parte_determination = "Appears Ex Parte" AND ex_parte_denial_explanation <> "" THEN err_msg = err_msg & vbCr & "* The explanation for denial field should be blank since ex parte has been approved."

				'Add validation that if ex parte denied, then explanation must be provided
				If ex_parte_determination = "Cannot be Processed as Ex Parte" AND ex_parte_denial_explanation = "" THEN err_msg = err_msg & vbCr & "* You must provide an explanation for the ex parte denial."

				If len(ex_parte_denial_explanation) > 255 Then err_msg = err_msg & vbCr & "* The explanation for the Ex Parte denial is too long and should be shortened. The length of the information cannot be more than 255 character."

				'Add validation to ensure worker signature is not blank
				IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please include your worker signature."

				'Error message handling
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = "" AND ButtonPressed = -1
			'Add to all dialogs where you need to work within BLUEZONE
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
		'End dialog section-----------------------------------------------------------------------------------------------


		'Checks to see if in MAXIS
		Call check_for_MAXIS(False)

		'Ensure starting at SELF so that writing to CASE NOTE works properly
		CALL back_to_SELF()

		If ex_parte_determination = "Appears Ex Parte" or ex_parte_determination = "Cannot be Processed as Ex Parte" Then

			'Navigate to STAT, REVW, and open HC Renewal Window with instructions
			CALL navigate_to_MAXIS_screen("STAT", "REVW")
			CALL write_value_and_transmit("X", 5, 71)

			'Read data from HC renewal screen to determine what changes the worker needs to complete and then use to validate changes
			'TO DO - update variables to match/pull from SQL data table. This data should be used as baseline/reference point for validation.
			EMReadScreen income_renewal_date, 8, 7, 27
			' EMReadScreen elig_renewal_date, 8, 8, 27
			elig_renewal_date = review_month_from_SQL
			EMReadScreen HC_ex_parte_determination, 1, 9, 27
			EMReadScreen income_asset_renewal_date, 8, 7, 71
			EMReadScreen exempt_6_mo_ir_form, 1, 8, 71
			EMReadScreen ex_parte_renewal_month_year, 7, 9, 71

			'Dialog and review of HC renewal for approval of ex parte
			If ex_parte_determination = "Appears Ex Parte" Then

				Dialog1 = "" 'blanking out dialog name

				BeginDialog Dialog1, 0, 0, 331, 150, "Health Care Renewal Updates - Appears Ex Parte"
				ButtonGroup ButtonPressed
					PushButton 205, 130, 100, 15, "Verify HC Renewal Updates", hc_renewal_button
				Text 5, 5, 320, 10, "Update the following on the Health Care Renewals Screen and then click the button below to verify:"
				Text 10, 20, 270, 10, "- Elig Renewal Date: Enter one year from the renewal month/year currently listed"
				Text 10, 35, 100, 10, "- Income/Asset Renewal Date:"
				Text 25, 45, 290, 20, "- For cases with a spenddown that do not meet an exception listed in EPM 2.3.4.2 MA-ABD Renewals, enter a date six months from the date updated in ELIG Renewal Date"
				Text 25, 65, 275, 10, "- For all other cases, enter the same date entered in the Elig Renewal Date"
				Text 10, 80, 145, 10, "- Exempt from 6 Mo IR: Enter N"
				Text 10, 95, 145, 10, "- ExParte: Enter Y"
				Text 10, 110, 255, 10, "- ExParte Renewal Month: Enter month and year of the ex parte renewal month"
				EndDialog


				DO
					Do
						err_msg = ""    'This is the error message handling
						Dialog Dialog1
						cancel_confirmation

						' If ButtonPressed = hc_renewal_button Then Call check_hc_renewal_updates() ' TO DO - timing of function calls and completing function within loop?

						'TO DO - update with functions?
						'Check the HC renewal screen data and compare against initial to ensure that changes made properly

						'TODO - read DISA and display waiver

						EMReadScreen stat_check, 4, 20, 21
						EMReadScreen revw_panel_check, 4, 2, 46
						EMReadScreen hc_revw_pop_up_check, 20, 4, 32

						If hc_revw_pop_up_check <> "HEALTH CARE RENEWALS" Then
							If hc_revw_pop_up_check = "REVW" Then
								EMReadScreen pop_up_open, 1, 4, 22
								If pop_up_open <> "*" Then PF3
								' Call write_value_and_transmit({"X", 5, 71)
							ElseIf stat_check = "STAT" Then
								Call write_value_and_transmit("REVW", 20, 71)
								' Call write_value_and_transmit({"X", 5, 71)
							Else
								Call MAXIS_background_check
								CALL navigate_to_MAXIS_screen("STAT", "REVW")
							End If
							CALL write_value_and_transmit("X", 5, 71)
						End If

						' CALL back_to_SELF()
						' CALL navigate_to_MAXIS_screen("STAT", "REVW")
						' CALL write_value_and_transmit("X", 5, 71)
						' EMReadScreen check_income_renewal_date, 8, 7, 27
						EMReadScreen check_elig_renewal_date, 8, 8, 27
						EMReadScreen check_HC_ex_parte_determination, 1, 9, 27
						EMReadScreen check_income_asset_renewal_date, 8, 7, 71
						If check_income_asset_renewal_date = "__ 01 __" Then EMReadScreen check_income_asset_renewal_date, 8, 7, 27
						EMReadScreen check_exempt_6_mo_ir_form, 1, 8, 71
						EMReadScreen check_ex_parte_renewal_month_year, 7, 9, 71

						check_elig_renewal_date = replace(check_elig_renewal_date, " ", "/")
						check_income_asset_renewal_date = replace(check_income_asset_renewal_date, " ", "/")
						elig_renewal_date = replace(elig_renewal_date, " ", "/")
						' income_asset_renewal_date = replace(income_asset_renewal_date, " ", "/")

						check_elig_renewal_date = DateAdd("d", 0, check_elig_renewal_date)
						check_income_asset_renewal_date = DateAdd("d", 0, check_income_asset_renewal_date)
						elig_renewal_date = DateAdd("d", 0, elig_renewal_date)
						' income_asset_renewal_date = DateAdd("d", 0, income_asset_renewal_date)

						' MsgBox "check_elig_renewal_date - " & check_elig_renewal_date & vbCr & "elig_renewal_date - " & elig_renewal_date
						'Validate Elig Renewal Date to ensure it is set for 1 year from current Elig Renewal Date
						If check_elig_renewal_date <> DateAdd("yyyy", 1, elig_renewal_date) THEN err_msg = err_msg & vbCr & "* The Elig Renewal Date should be set for 1 year from the current renewal month and year."

						'Validate Income/Asset Renewal Date to ensure it is the same as the Elig Renewal Date or set for 6 months from original Elig Renewal Date for cases with a spenddown:
						'TO DO - determine how to determine if meets spenddown?
						' If check_income_asset_renewal_date <> DateAdd("Y", 1, income_asset_renewal_date) OR check_income_asset_renewal_date <> DateAdd("M", 6, income_asset_renewal_date) THEN
						If check_income_asset_renewal_date <> check_elig_renewal_date Then err_msg = err_msg & vbCr & "* The Income/Asset Renewal Date should be be the same as the Elig Renewal Date. For cases with a spenddown that do not meet an exception listed in EPM 2.3.4.2 MA-ABD Renewals, enter a date six months from the original ELIG Renewal Date."
						'TODO - put back the funcitonality for spenddowns

						'Validate that Exempt from 6 Mo IR is set to N
						If check_exempt_6_mo_ir_form <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for Exempt from 6 Mo IR."

						'Validate that ExParte field updated to Y
						If check_HC_ex_parte_determination <> "Y" THEN err_msg = err_msg & vbCr & "* You must enter 'Y' for ExParte."

						'Validate that ExParte Renewal Month is correct
						'TO DO - add validation to ensure that date updated in HC renewal screen is the same as date provided in SQL table
						If check_ex_parte_renewal_month_year = "__ ____" THEN err_msg = err_msg & vbCr & "* You must enter the month and year for the Ex Parte renewal month."
						If check_ex_parte_renewal_month_year <> correct_ex_parte_revw_month_code Then err_msg = err_msg & vbCr & "* The ExParte Renewal Month on REVW should be " & correct_ex_parte_revw_month_code & "."

						'Error message handling
						IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
					Loop until err_msg = ""
						'Add to all dialogs where you need to work within BLUEZONE
						CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
				LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
			End If

			'Dialog and review of HC renewal for denial of ex parte
			If ex_parte_determination = "Cannot be Processed as Ex Parte" Then

				Dialog1 = ""

				BeginDialog Dialog1, 0, 0, 331, 120, "Health Care Renewal Updates - Cannot be Processed as Ex Parte"
				ButtonGroup ButtonPressed
					PushButton 205, 100, 100, 15, "Verify HC Renewal Updates", hc_renewal_button
				Text 5, 5, 320, 10, "Update the following on the Health Care Renewals Screen and then click the button below to verify:"
				Text 10, 20, 150, 10, "- Elig Renewal Date: Should not be changed"
				Text 10, 35, 300, 10, "- Income/Asset Renewal Date: Should not be changed and should match Elig Renewal Date."
				Text 10, 50, 145, 10, "- Exempt from 6 Mo IR: Enter N"
				Text 10, 65, 145, 10, "- ExParte: Enter N"
				Text 10, 80, 255, 10, "- ExParte Renewal Month: Enter month and year of the ex parte renewal month"
				EndDialog


				DO
					Do
						err_msg = ""    'This is the error message handling
						Dialog Dialog1
						cancel_confirmation

						' If ButtonPressed = hc_renewal_button Then Call check_hc_renewal_updates() ' TO DO - timing of function calls and completing function within loop?

						'TO DO - update with functions?
						'Check the HC renewal screen data and compare against initial to ensure that changes made properly

						EMReadScreen stat_check, 4, 20, 21
						EMReadScreen revw_panel_check, 4, 2, 46
						EMReadScreen hc_revw_pop_up_check, 20, 4, 32

						If hc_revw_pop_up_check <> "HEALTH CARE RENEWALS" Then
							If hc_revw_pop_up_check = "REVW" Then
								EMReadScreen pop_up_open, 1, 4, 22
								If pop_up_open <> "*" Then PF3
								' Call write_value_and_transmit({"X", 5, 71)
							ElseIf stat_check = "STAT" Then
								Call write_value_and_transmit("REVW", 20, 71)
								' Call write_value_and_transmit({"X", 5, 71)
							Else
								Call MAXIS_background_check
								CALL navigate_to_MAXIS_screen("STAT", "REVW")
							End If
							CALL write_value_and_transmit("X", 5, 71)
						End If

						EMReadScreen check_income_renewal_date, 8, 7, 27
						EMReadScreen check_elig_renewal_date, 8, 8, 27
						EMReadScreen check_HC_ex_parte_determination, 1, 9, 27
						EMReadScreen check_income_asset_renewal_date, 8, 7, 71
						If check_income_asset_renewal_date = "__ 01 __" Then EMReadScreen check_income_asset_renewal_date, 8, 7, 27
						EMReadScreen check_exempt_6_mo_ir_form, 1, 8, 71
						EMReadScreen check_ex_parte_renewal_month_year, 7, 9, 71

						check_elig_renewal_date = replace(check_elig_renewal_date, " ", "/")
						check_income_asset_renewal_date = replace(check_income_asset_renewal_date, " ", "/")
						elig_renewal_date = replace(elig_renewal_date, " ", "/")
						' income_asset_renewal_date = replace(income_asset_renewal_date, " ", "/")

						check_elig_renewal_date = DateAdd("d", 0, check_elig_renewal_date)
						check_income_asset_renewal_date = DateAdd("d", 0, check_income_asset_renewal_date)
						elig_renewal_date = DateAdd("d", 0, elig_renewal_date)
						' income_asset_renewal_date = DateAdd("d", 0, income_asset_renewal_date)

						'Validation to ensure that elig renewal date has not changed
						If check_elig_renewal_date <> elig_renewal_date THEN err_msg = err_msg & vbCr & "* The Elig Renewal Date should not have been changed. It should remain " & elig_renewal_date & "."

						'Validation for Income/Asset Renewal Date to ensure that information has not changed
						If check_income_asset_renewal_date <> check_elig_renewal_date THEN err_msg = err_msg & vbCr & "* The Income/Asset Renewal Date should not have been changed. It should remain " & income_asset_renewal_date & "."

						'Validation to ensure that Exempt from 6 Mo IR is set to N
						If check_exempt_6_mo_ir_form <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for Exempt from 6 Mo IR."

						'Validation to ensure that ExParte field updated to N
						If check_HC_ex_parte_determination <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for ExParte."

						'Validate that ExParte Renewal Month is correct
						'TO DO - add validation to ensure that date updated in HC renewal screen is the same as date provided in SQL table
						If check_ex_parte_renewal_month = "__ ____" THEN err_msg = err_msg & vbCr & "* You must enter the month and year for the Ex Parte renewal month."
						If check_ex_parte_renewal_month_year <> correct_ex_parte_revw_month_code Then err_msg = err_msg & vbCr & "* The ExParte Renewal Month on REVW should be " & correct_ex_parte_revw_month_code & "."

						'Error message handling
						IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
					Loop until err_msg = ""
						'Add to all dialogs where you need to work within BLUEZONE
						CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
				LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
			End If
		End If
		' MsgBox "STOP HERE FOR NOW!!!"

		If ex_parte_determination = "Appears Ex Parte" Then appears_ex_parte = True
		If ex_parte_determination = "Cannot be Processed as Ex Parte" Then appears_ex_parte = False
		If ex_parte_determination = "Health Care has been Closed" Then appears_ex_parte = False
		If ex_parte_determination = "Case Transfered Out of County" Then appears_ex_parte = False
		' MsgBox "appears_ex_parte - " & appears_ex_parte

		If MX_region <> "TRAINING" Then
			If user_ID_for_validation <> "CALO001" AND user_ID_for_validation <> "MARI001" Then
				' MsgBox "STOP - YOU ARE GOING TO UPDATE"
				sql_format_ex_parte_denial_explanation = replace(ex_parte_denial_explanation, "'", "")
				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', Phase1HSR = '" & user_ID_for_validation & "', ExParteAfterPhase1 = '" & ex_parte_determination & "', Phase1ExParteCancelReason = '" & sql_format_ex_parte_denial_explanation & "' WHERE CaseNumber = '" & SQL_Case_Number & "'"

				'Creating objects for Access
				Set objUpdateConnection = CreateObject("ADODB.Connection")
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'This is the file path for the statistics Access database.
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else
				MsgBox "This is where the update would happen" & vbCr & vbCr & "appears_ex_parte - " & appears_ex_parte & vbCr& "user_ID_for_validation - " & user_ID_for_validation & vbCr & "ex_parte_determination - " & ex_parte_determination & vbCr & "ex_parte_denial_explanation - " & ex_parte_denial_explanation
			End If
		End If
		' MsgBox "About to CASE NOTE AND TIKL"
		'If ex parte approved, create TIKL for 1st of processing month which is renewal month - 1
		If ex_parte_determination = "Appears Ex Parte" Then Call create_TIKL("Phase 1 - The case has been evaluated for ex parte and appears to be ex parte on the information provided.", 0, DateAdd("M", -1, elig_renewal_date), False, TIKL_note_text)

		If ex_parte_determination = "Case Transfered Out of County" Then script_end_procedure("Case List updated with Ex parte Evaluation. No CASE/NOTE as case is not in Hennepin County.")
		'Navigate to and start a new CASE NOTE
		Call start_a_blank_case_note

		'Add title to CASE NOTE
		CALL write_variable_in_case_note("*** EX PARTE DETERMINATION - " & UCASE(ex_parte_determination) & " ***")

		'For ex parte approval, write information to case note
		If ex_parte_determination = "Appears Ex Parte" Then
			CALL write_variable_in_case_note(TIKL_note_text)
			CALL write_variable_in_case_note("Phase 1 - The case has been evaluated for ex parte and appears to be ex parte on the current case information.")
			CALL write_variable_in_case_note("MA-ABD enrollees will be Ex Parte renewed if their income can be verified electronically without the need for residents to provide verifications.")
		End If


		'For ex parte denial, write information to case note
		If ex_parte_determination = "Cannot be Processed as Ex Parte" Then
			CALL write_variable_in_case_note("Phase 1 - The case has been evaluated for ex parte and cannot be processed as Ex Parte Renewal based on the information provided.")
			CALL write_bullet_and_variable_in_case_note("Reason for Denial", ex_parte_denial_explanation)
		End If

		If ex_parte_determination = "Health Care has been Closed" Then
			CALL write_variable_in_case_note("No renewal required on case as the Health Care is no longer active.")
		End If

		'Add worker signature
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)

		'Script end procedure
		script_end_procedure("Success! The ex parte review information has been added to the CASE NOTE")
	End If

	If ex_parte_phase = "Phase 2" Then
		Call convert_date_into_MAXIS_footer_month(review_month_from_SQL, ex_parte_renewal_month, ex_parte_renewal_year)

		updated_hc_renewal_month = CM_plus_2_mo
		updated_hc_renewal_year = CM_plus_2_yr

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 351, 230, "Phase 2 - Ex Parte Denied"
			DropListBox 15, 75, 320, 45, "Select One..."+chr(9)+"Reschedule Renewal - HC Eligibility does not maintain ongoing."+chr(9)+"Case closed during the blackout period."+chr(9)+"Case transfered out of Hennepin during the blackout period.", phase_2_denial_reason
			EditBox 15, 105, 320, 15, phase_1_changes_summary
			EditBox 15, 135, 320, 15, phase_2_notes
			EditBox 215, 175, 20, 15, updated_hc_renewal_month
			EditBox 240, 175, 20, 15, updated_hc_renewal_year
			EditBox 80, 205, 125, 15, worker_signature
			ButtonGroup ButtonPressed
				OkButton 235, 205, 50, 15
				CancelButton 290, 205, 50, 15
			GroupBox 10, 5, 330, 35, "Case Info"
			Text 15, 20, 20, 10, "Case:"
			Text 50, 20, 80, 10, MAXIS_case_number
			Text 155, 20, 145, 15, "This script is only used for cases where an Ex Parte Renewal CANNOT be approved."
			GroupBox 10, 50, 330, 105, "Ex Parte Denial Explanation"
			Text 15, 65, 285, 10, "Reason Ex Parte cannot be approved:"
			Text 15, 95, 225, 10, "Explain what changed from the Evaluation to prevent approval:"
			Text 15, 125, 90, 10, "Additional CASE/NOTEs:"
			GroupBox 10, 165, 330, 30, "Update HC Renewal Date"
			Text 15, 180, 190, 10, "What month is the standard Renewal being scheduled for?"
			Text 15, 210, 60, 10, "Worker Signature:"
		EndDialog

		DO
			Do
				err_msg = ""    'This is the error message handling
				Dialog Dialog1
				cancel_confirmation

				'Trim information for CASE/NOTES
				phase_1_changes_summary = trim(phase_1_changes_summary)
				phase_2_notes = trim(phase_2_notes)

				If phase_2_denial_reason = "Select One..." Then err_msg = err_msg & vbCr & "* You must select the reason this case cannot be processed as an Ex Parte Approval."
				If phase_2_denial_reason = "Reschedule Renewal - HC Eligibility does not maintain ongoing." Then
					If phase_1_changes_summary = "" Then err_msg = err_msg & vbCr & "* You must provide a summary of any changes since the Phase 1 determination for this case."
					If len(phase_1_changes_summary) > 255 Then err_msg = err_msg & vbCr & "* The changes during the blackout period denial is too long and should be shortened. The length of the information cannot be more than 255 characters."
					Call validate_footer_month_entry(updated_hc_renewal_month, updated_hc_renewal_year, err_msg, "*")
				End If

				IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please include your worker signature."

				'Error message handling
				IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
			Loop until err_msg = ""
			'Add to all dialogs where you need to work within BLUEZONE
			CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		If phase_2_denial_reason = "Reschedule Renewal - HC Eligibility does not maintain ongoing." Then ex_parte_after_phase_2 = updated_hc_renewal_month & "/" & updated_hc_renewal_year & " ER Scheduled"
		If phase_2_denial_reason = "Case closed during the blackout period." Then ex_parte_after_phase_2 = "Closed HC"
		If phase_2_denial_reason = "Case transfered out of Hennepin during the blackout period." Then ex_parte_after_phase_2 = "Case not in 27"


		'Checks to see if in MAXIS
		Call check_for_MAXIS(False)

		'Ensure starting at SELF so that writing to CASE NOTE works properly
		CALL back_to_SELF()

		If phase_2_denial_reason = "Reschedule Renewal - HC Eligibility does not maintain ongoing." Then

			'Navigate to STAT, REVW, and open HC Renewal Window with instructions
			CALL navigate_to_MAXIS_screen("STAT", "REVW")
			CALL write_value_and_transmit("X", 5, 71)

			'Open dialog to verify changes to HC renewal screen
			Dialog1 = "" 'blanking out dialog name
			BeginDialog Dialog1, 0, 0, 301, 175, "STAT/REVW Coding to Reschedule Standard Renewal"
				ButtonGroup ButtonPressed
					PushButton 190, 150, 100, 15, "Verify HC Renewal Updates", hc_renewal_button
				Text 10, 10, 230, 20, "To reschedule the Renewal Month for a standard Renewal, the STAT/REVW panel should be reviewed and updated correctly."
				Text 10, 40, 210, 10, "This is what the STAT/REVW HC Pop-Up should look like:"
				GroupBox 10, 55, 280, 55, "HEALTH CARE RENEWALS"
				Text 20, 70, 110, 10, "Income Renewal Date: " & updated_hc_renewal_month & " 01 " & updated_hc_renewal_year
				Text 150, 70, 125, 10, "Income/Asset Renewal Date: __ __ __"
				Text 30, 80, 110, 10, "Elig Renewal Date: " & updated_hc_renewal_month & " 01 " & updated_hc_renewal_year
				Text 175, 80, 85, 10, "Exempt from 6 Mo IR: N"
				Text 45, 90, 60, 10, "ExParte (Y/N): N"
				Text 165, 90, 120, 10, "ExParte Renewal Month: " & ex_parte_renewal_month & " 20" & ex_parte_renewal_year
				Text 15, 120, 265, 10, "If these fields are different, CHANGE THEM NOW, while this dialog is displayed."
				Text 15, 130, 200, 10, "(IR and AR coding can be switched based on case scenario.)"
			EndDialog

			DO
				Do
					err_msg = ""    'This is the error message handling
					Dialog Dialog1
					cancel_confirmation

					EMReadScreen stat_check, 4, 20, 21
					EMReadScreen revw_panel_check, 4, 2, 46
					EMReadScreen hc_revw_pop_up_check, 20, 4, 32

					If hc_revw_pop_up_check <> "HEALTH CARE RENEWALS" Then
						If hc_revw_pop_up_check = "REVW" Then
							EMReadScreen pop_up_open, 1, 4, 22
							If pop_up_open <> "*" Then PF3
							' Call write_value_and_transmit({"X", 5, 71)
						ElseIf stat_check = "STAT" Then
							Call write_value_and_transmit("REVW", 20, 71)
							' Call write_value_and_transmit({"X", 5, 71)
						Else
							Call MAXIS_background_check
							CALL navigate_to_MAXIS_screen("STAT", "REVW")
						End If
						CALL write_value_and_transmit("X", 5, 71)
					End If

					'Read data from HC renewal screen to determine what changes the worker needs to complete and then use to validate changes
					EMReadScreen income_renewal_date, 8, 7, 27
					EMReadScreen income_asset_renewal_date, 8, 7, 71
					If income_renewal_date <> "__ __ __" Then
						sr_date = income_renewal_date
					ElseIf	income_asset_renewal_date <> "__ __ __" Then
						sr_date = income_asset_renewal_date
					End If
					sr_month = left(sr_date, 2)
					sr_year = right(sr_date, 2)
					EMReadScreen elig_renewal_month, 2, 8, 27
					EMReadScreen elig_renewal_year, 2, 8, 33
					EMReadScreen exempt_6_mo_ir_form, 1, 8, 71
					EMReadScreen HC_ex_parte_determination, 1, 9, 27
					EMReadScreen REVW_ex_parte_renewal_month, 2, 9, 71
					EMReadScreen REVW_ex_parte_renewal_year, 2, 9, 76

					revw_panel_err = ""

					If elig_renewal_month <> updated_hc_renewal_month or elig_renewal_year <> updated_hc_renewal_year Then
						revw_panel_err = revw_panel_err & vbCr & "* The 'Elig Renewal Date' needs to be coded for the month you entered as the next available month for a standard renewal could be completed. You listed this as " & updated_hc_renewal_month & "/" & updated_hc_renewal_year & "."
						If elig_renewal_month <> updated_hc_renewal_month Then revw_panel_err = revw_panel_err & vbCr & "* The month is incorrect, the panel has " & elig_renewal_month & " listed as the month."
						If elig_renewal_year <> updated_hc_renewal_year Then revw_panel_err = revw_panel_err & vbCr & "* The year is incorrect, the panel has " & elig_renewal_year & " listed as the year."
					End If

					If sr_month <> updated_hc_renewal_month or sr_year <> updated_hc_renewal_year Then
						revw_panel_err = revw_panel_err & vbCr & "* The Six Month Report (IR/AR) should be aligned with the correct Elig Renewal Month that you entered as the next available month for a standard renewal (" & updated_hc_renewal_month & "/" & updated_hc_renewal_year & ")"
						If sr_month <> updated_hc_renewal_month Then revw_panel_err = revw_panel_err & vbCr & "* The month for the IR/AR is incorrect, the panel has " & sr_month & " listed as the month."
						If sr_year <> updated_hc_renewal_year Then revw_panel_err = revw_panel_err & vbCr & "* The month for the IR/AR is incorrect, the panel has " & sr_year & " listed as the month."
					End If

					If HC_ex_parte_determination = "Y" Then revw_panel_err = revw_panel_err & vbCr & "* Update the Ex Parte yes/no indicator to 'N', since you are recording that you cannot process as Ex Parte."

					If REVW_ex_parte_renewal_month <> ex_parte_renewal_month or REVW_ex_parte_renewal_year <> ex_parte_renewal_year Then
						revw_panel_err = revw_panel_err & vbCr & "* The Ex Parte Renewal month should be left coded with the month that was evaluated (" & ex_parte_renewal_month & "/" & ex_parte_renewal_year & ") for recording the Ex Parte work."
						If REVW_ex_parte_renewal_month <> ex_parte_renewal_month Then revw_panel_err = revw_panel_err & vbCr & "* The Ex Parte month is incorrect, the panel has " & REVW_ex_parte_renewal_month & " entered."
						If REVW_ex_parte_renewal_year <> ex_parte_renewal_year Then revw_panel_err = revw_panel_err & vbCr & "* The Ex Parte month is incorrect, the panel has 20" & REVW_ex_parte_renewal_year & " entered."
					End If

					'Error message handling
					IF revw_panel_err <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & revw_panel_err & vbNewLine
				Loop until revw_panel_err = ""
				CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
			LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
		End If

		'TO DO - verify that ex parte determination update to SQL database is correct
		ex_parte_determination = "Cannot be Processed as Ex Parte"

		If MX_region <> "TRAINING" Then
			If user_ID_for_validation <> "CALO001" AND user_ID_for_validation <> "MARI001" Then
				sql_format_phase_2_denial_reason = replace(phase_1_changes_summary, "'", "")
				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2HSR = '" & user_ID_for_validation & "', ExParteAfterPhase2 = '" & ex_parte_after_phase_2 & "', Phase2ExParteCancelReason = '" & sql_format_phase_2_denial_reason & "' WHERE CaseNumber = '" & SQL_Case_Number & "'"

				'Creating objects for Access
				Set objUpdateConnection = CreateObject("ADODB.Connection")
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'This is the file path for the statistics Access database.
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else
				MsgBox "This is where the update would happen" & vbCr & vbCr & "user_ID_for_validation - " & user_ID_for_validation & vbCr & "ex_parte_after_phase_2 - " & ex_parte_after_phase_2 & vbCr & "phase_1_changes_summary - " & phase_1_changes_summary
			End If
		End If

		If phase_2_denial_reason = "Reschedule Renewal - HC Eligibility does not maintain ongoing." Then
			Call start_a_blank_CASE_NOTE
			Call write_variable_in_CASE_NOTE(updated_hc_renewal_month & "/" & updated_hc_renewal_year & " HC ER Scheduled - Ex Parte Could not be Processed")
			Call write_variable_in_CASE_NOTE("Unable to complete all necessary Verifications to continue HC.")
			Call write_variable_in_CASE_NOTE("Health Care Enrollees will need to complete a standard Renewal for " & updated_hc_renewal_month & "/" & updated_hc_renewal_year)
			Call write_variable_in_CASE_NOTE(updated_hc_renewal_month & "/" & updated_hc_renewal_year & " is the next available month for and Eligibility Review.")
			Call write_variable_in_CASE_NOTE("---")
			CALL write_bullet_and_variable_in_case_note("Changes since Phase 1", phase_1_changes_summary)
			CALL write_bullet_and_variable_in_case_note("Additional notes", phase_2_notes)
			Call write_variable_in_CASE_NOTE("---")
			Call write_variable_in_CASE_NOTE(worker_signature)
		End If

		If phase_2_denial_reason = "Case closed during the blackout period." Then
			Call start_a_blank_CASE_NOTE
			CALL write_variable_in_case_note("Ex Parte Could not be Processed as HC was Closed")
			CALL write_variable_in_case_note("Phase 2 - The case has been evaluated for ex parte but was closed before the Ex Parte Approval was completed.")
			CALL write_bullet_and_variable_in_case_note("Changes since Phase 1", phase_1_changes_summary)
			CALL write_bullet_and_variable_in_case_note("Additional notes", phase_2_notes)
			CALL write_variable_in_case_note("---")
			CALL write_variable_in_case_note(worker_signature)
		End If

		'Script end procedure
		script_end_procedure("Script run complete. Data table has been updated and CASE/NOTE created if necessary. Ex Parte Phase 2 could not be completed for this case.")
	End If
End If

'determing if the application date is before or after 4/1/23
applied_after_03_23 = True
cutoff_date = #4/1/2023#
If DateDiff("d", form_date, cutoff_date) > 0 Then applied_after_03_23 = False

'Read PROG and HCRE to gather appliation date and any retro request
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "PROG", is_this_priv)
If is_this_priv = True Then Call script_end_procedure("This case appears PRIVILEGED. The script will now end as there is no access to case information.")
EMReadScreen case_county, 4, 21, 21
If case_county <> worker_county_code Then Call script_end_procedure("This case does not appear to be in your county and there is no access to case action. The script will now end.")
EMReadScreen prog_hc_appl_date, 8, 12, 33
EMReadScreen prog_hc_intvw_date, 8, 12, 55
EMReadScreen prog_hc_status, 4, 12, 74

'creating a list of all the HH members for the dialog dropdown
Call generate_client_list(case_memb_list, "Select or Type Member")
verification_memb_list = " "+chr(9)+case_memb_list

If prog_hc_appl_date = "__ __ __" Then prog_hc_appl_date = ""			'formatting the date information
prog_hc_appl_date = replace(prog_hc_appl_date, " ", "/")
If prog_hc_intvw_date = "__ __ __" Then prog_hc_intvw_date = ""
prog_hc_intvw_date = replace(prog_hc_intvw_date, " ", "/")
hc_application_date = prog_hc_appl_date

If prog_hc_status = "PEND" Then health_care_pending = True			'determine if we have HC Pending
If prog_hc_status = "ACTV" Then health_care_active = True
'TODO - add better handling for REVW

Call navigate_to_MAXIS_screen("STAT", "HCRE")						'going to read who is listed on HCRE
hc_memb = 0
hc_row = 10
Do
	EMReadScreen hcre_ref_numb, 2, hc_row, 24
	If hcre_ref_numb <> "  " Then
		ReDim Preserve HEALTH_CARE_MEMBERS(last_const, hc_memb)

		HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) = hcre_ref_numb
		HEALTH_CARE_MEMBERS(pers_btn_one_const, hc_memb) = 500+hc_memb

		If HC_form_name = "Breast and Cervical Cancer Coverage Group (DHS-3525)" Then
			HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, hc_memb) = "Breast/Cervical Cancer"
			HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, hc_memb) = "No Eligibility"
		End If

		EMReadScreen HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb), 8, hc_row, 51
		EMReadScreen hc_appl_mo, 2, hc_row, 51
		EMReadScreen hc_appl_yr, 2, hc_row, 57
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_mo_const, hc_memb), 2, hc_row, 64
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_yr_const, hc_memb), 2, hc_row, 67
		EMReadScreen HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb), 5, hc_row, 64

		HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb) = replace(HEALTH_CARE_MEMBERS(hc_cov_date_const, hc_memb), " ", "/")
		HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = False
		HEALTH_CARE_MEMBERS(member_is_recert_for_hc_const, hc_memb) = False
		HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = False
		If hc_appl_mo <> HEALTH_CARE_MEMBERS(hc_cov_mo_const, hc_memb) or hc_appl_yr <> HEALTH_CARE_MEMBERS(hc_cov_yr_const, hc_memb) Then HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = True
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

'Now we go read STAT/MEMB for all of the persons listed on HCRE
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

'reading from CASE/CURR to get the case information
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
curr_hc_membs = " "
all_membs_with_hcre = " "
Do
	EMReadScreen pers_memb_numb, 2, pers_row, 3
	EMReadScreen pers_hc_status, 1, pers_row, 61

	For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
		' MsgBox "pers_memb_numb - " & pers_memb_numb & vbCr & "HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) - " & HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb)
		If pers_memb_numb = HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) Then
			HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) = pers_hc_status
			all_membs_with_hcre = all_membs_with_hcre & HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) & " "
			' MsgBox "HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) - " & HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) & " - 1"
			If pers_hc_status = "I" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "INACTIVE"
			If pers_hc_status = "D" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "DENIED"
			If pers_hc_status = "A" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "ACTIVE"
			If pers_hc_status = "P" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "PENDING"
			If pers_hc_status = "R" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = "REINSTATEMENT"
			' If pers_hc_status = "" Then HEALTH_CARE_MEMBERS(case_pers_hc_status_info_const, hc_memb) = ""
			If pers_hc_status = "P" Then health_care_pending = True
			If pers_hc_status = "A" Then health_care_active = True
			If pers_hc_status = "A" or pers_hc_status = "R" or pers_hc_status = "P" Then curr_hc_membs = curr_hc_membs & HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) & " "

			' MsgBox "hc_application_date - " & hc_application_date & vbCr & "HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) - " & HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)
			If DateDiff("d", hc_application_date, HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)) > 0 Then
				hc_application_date = HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb)
			End If
		End If
	Next

	If pers_memb_numb = "  " Then Exit Do

	pers_row = pers_row + 3
	If pers_row = 19 Then
		pers_row = 10
		PF8
		EMReadScreen last_page_check, 9, 24, 14
	End If
Loop until last_page_check = "LAST PAGE"
curr_hc_membs = trim(curr_hc_membs)
all_membs_with_hcre = trim(all_membs_with_hcre)

case_has_retro_request = False
For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)					'setting the defaults for booleans for each member with HC
	HEALTH_CARE_MEMBERS(show_hc_detail_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(DISA_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(PREG_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(PARE_exists_const, hc_memb) = False
	HEALTH_CARE_MEMBERS(MEDI_exists_const, hc_memb) = False

	If hc_application_date = HEALTH_CARE_MEMBERS(hc_appl_date_const, hc_memb) Then HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True
	If HEALTH_CARE_MEMBERS(case_pers_hc_status_code_const, hc_memb) = "P" Then HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True
	HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True			'TODO - remove this when I can figure out who is actually requesting vs recertifying

	If HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = True Then
		call read_person_based_STAT_info()
	End If
	If HEALTH_CARE_MEMBERS(member_is_applying_for_hc_const, hc_memb) = False Then HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = False
Next

'this is special handling for Presumptive Eligibility for MA-BC - whcih is processed off of these two forms.
If HC_form_name = "SAGE Enrollment Form (MA/BC PE Only)" or HC_form_name = "Screen Our Circle Form (MA/BC PE Only)" Then
	first_month_pe = form_date									'determining the months and other dates for MA-BC PE based on the form date
	next_month_pe = DateAdd("m", 1, form_date)
	first_mo_mo = right("00" & DatePart("m", first_month_pe), 2)
	first_mo_yr = right(DatePart("yyyy", first_month_pe), 2)
	second_mo_mo = right("00" & DatePart("m", next_month_pe), 2)
	second_mo_yr = right(DatePart("yyyy", next_month_pe), 2)
	first_month_pe = first_mo_mo & "/" & first_mo_yr
	next_month_pe = second_mo_mo & "/" & second_mo_yr
	temp_ma_auth_form_date = form_date
	end_pe_tikl_date = second_mo_mo & "/1/" & second_mo_yr
	end_pe_tikl_date = DateAdd("d", 0, end_pe_tikl_date)

	'special MA-BC PE Eligibility
	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 381, 235, "MA-BC Presumptive Eligiblity"
  			  ButtonGroup ButtonPressed
				DropListBox 195, 10, 180, 45, case_memb_list, ma_bc_member
				EditBox 310, 30, 65, 15, temp_ma_auth_form_date
				CheckBox 10, 130, 325, 10, "Check here to have the script set a TIKL for " & end_pe_tikl_date & " to close MA-BC Presumptive Eligibliity.", tikl_to_close_PE_checkbox
				EditBox 10, 160, 365, 15, ma_bc_notes
				EditBox 10, 190, 365, 15, ma_bc_tikls
				OkButton 270, 215, 50, 15
				CancelButton 325, 215, 50, 15
				Text 10, 15, 185, 10, "Select the Person with MA/BC Presumptive Eligibility:"
				Text 10, 35, 300, 10, "Enter the date the Temporary Medical Assistance Authorization (DHS-3525B) was received:"
				Text 10, 55, 350, 10, HC_form_name & " received on " & form_date
				Text 10, 70, 115, 10, "Case Information for CASE/NOTE:"
				Text 20, 85, 315, 10, "Breast Cancer application if Health Care is still needed after 2 months of Presumptive Care."
				Text 20, 100, 250, 10, "Method X Budget - no Income or Assets needed for Presumptive Eligibility."
				Text 20, 115, 205, 10, "Presumptive Care to be approved for " & first_month_pe &"  and " & next_month_pe & "."
				Text 10, 150, 85, 10, "Additional Case Details:"
				Text 10, 180, 85, 10, "Additional TIKLs Set"
			EndDialog

			Dialog Dialog1
			cancel_confirmation

			ma_bc_member = trim(ma_bc_member)
			ma_bc_notes = trim(ma_bc_notes)
			ma_bc_tikls = trim(ma_bc_tikls)
			temp_ma_auth_form_date = trim(temp_ma_auth_form_date)

			If ma_bc_member = "" or ma_bc_member = "Select One..." Then err_msg = err_msg & vbCr & "* Select the Household member who is receiving MA-BC PE."
			If IsDate(temp_ma_auth_form_date) = False Then err_msg = err_msg & vbCr & "* Enter the date the Copy of the Temporaty Medical Assistance Authorization (DHS-3525B) was received."

			If err_msg <> "" Then MsgBox "* * * * * * NOTICE * * * * * *" & vbCr & vbCr & "Please Resolve to Continue:" & vbCr & err_msg

		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False

	'setting a TIKL if selected
	If tikl_to_close_PE_checkbox = checked then Call create_TIKL("MA-BC Presumptive Eligibility ending " & second_mo_mo & "/" & second_mo_yr & ". Close case with 10-day notice.", 0, end_pe_tikl_date, False, MA_BC_PE_end_TIKL_note_text)

	'Entering the CASE/NOTE for MA-BC PE information
	short_form = replace(HC_form_name, "(MA/BC PE Only)", "")
	Call start_a_blank_case_note
	CALL write_variable_in_case_note(form_date & " " & short_form & ": Complete")

	Call write_bullet_and_variable_in_CASE_NOTE("Form Recvd", HC_form_name)
	If ltc_waiver_request_yn = "Yes" Then Call write_variable_in_CASE_NOTE("* This application can be used to request LTC/Waiver services.")
	Call write_bullet_and_variable_in_CASE_NOTE("Date Recvd", form_date)
	Call write_variable_in_CASE_NOTE("* Temporary Medical Assistance Authorization (DHS-3525B) recvd on: " & temp_ma_auth_form_date)
	Call write_variable_in_CASE_NOTE("========================== PERSON DETAILS ==========================")

	Call write_variable_in_CASE_NOTE("MEMB " & ma_bc_member & " approved for MA-BC Presumptive Care.")
	Call write_variable_in_CASE_NOTE("  Presumptive Care to be approved for " & first_month_pe & " and " & next_month_pe & ".")
	Call write_variable_in_CASE_NOTE("* Method X Budget - no Income or Assets needed for Presumptive Eligibility.")
	Call write_variable_in_CASE_NOTE("* Identity verified using medical document - " & short_form & ".")
	Call write_variable_in_CASE_NOTE("* Citizenship and Immigration information are not requested or required.")

	If ma_bc_notes <> "" OR ma_bc_tikls <> "" OR MA_BC_PE_end_TIKL_note_text <> "" Then Call write_variable_in_CASE_NOTE("============================== NOTES ===============================")
	Call write_bullet_and_variable_in_CASE_NOTE("Notes", ma_bc_notes)
	MA_BC_PE_end_TIKL_note_text = replace(MA_BC_PE_end_TIKL_note_text, ", 0 day return", "")
	Call write_variable_in_case_note(MA_BC_PE_end_TIKL_note_text & " TIKL to close MA-BC Presumptive Eligibility.")
	Call write_bullet_and_variable_in_CASE_NOTE("Additional TIKLs", ma_bc_tikls)

	Call write_variable_in_case_note("---")
	Call write_variable_in_case_note(worker_signature)

	Call script_end_procedure_with_error_report("MA-BC Presumptive Eligibility CASE/NOTE Created.")
End If

'gather information
'this is in place of the funtion -  HH_member_custom_dialog(HH_member_array)
'we need to check for count and process seperately.
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
transmit

DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
	'MsgBox access_denied_check
	If access_denied_check = "ACCESS DENIED" Then
		PF10
		EMWaitReady 0, 0
		last_name = "UNABLE TO FIND"
		first_name = " - Access Denied"
		mid_initial = ""
	Else
		EMReadscreen last_name, 25, 6, 30
		EMReadscreen first_name, 12, 6, 63
		EMReadscreen mid_initial, 1, 6, 79
		last_name = trim(replace(last_name, "_", "")) & " "
		first_name = trim(replace(first_name, "_", "")) & " "
		mid_initial = replace(mid_initial, "_", "")
	End If
	client_string = ref_nbr & last_name & first_name & mid_initial
	client_array = client_array & client_string & "|"
	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

If right(client_array, 1) = "|" Then client_array = left(client_array, len(client_array)-1)
client_array = TRIM(client_array)
test_array = split(client_array, "|")
total_clients = Ubound(test_array)			'setting the upper bound for how many spaces to use from the array

count_checkbox = 1
process_checkbox = 2
DIm all_clients_array()
ReDim all_clients_array(total_clients, 2)
Interim_array = split(client_array, "|")

If total_clients = 0 Then
	all_clients_array(0, 0) = Interim_array(0)
	all_clients_array(0, 1) = checked
	all_clients_array(0, 2) = checked
Else
	FOR x = 0 to total_clients				'using a dummy array to build in the autofilled check boxes into the array used for the dialog.
		all_clients_array(x, 0) = Interim_array(x)
		all_clients_array(x, 1) = checked
		ref_numb = left(Interim_array(x),2)
		If InStr(curr_hc_membs, ref_numb) <> 0 Then all_clients_array(x, 2) = checked
	NEXT

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 360, (85 + ((total_clients+1) * 15)), "HH Member Dialog"   'Creates the dynamic dialog. The height will change based on the number of clients it finds.
		Text 10, 5, 205, 10, "Select Household Members to capture information about."
		Text 10, 15, 205, 10, "Check all members: "
		Text 10, 25, 350, 10, "- In 'Count Income/Assets if their income or assets deem to anyone you are processing Health Care for."
		Text 10, 35, 350, 10, "- In 'Processing Health Care' if you are working on their Health Care Application or Renewal."
		Text 10, 55, 100, 10, "Count Income/Assets"
		Text 200, 55, 100, 10, "Processing Health Care"
		FOR i = 0 to total_clients										'For each person/string in the first level of the array the script will create a checkbox for them with height dependant on their order read
			IF all_clients_array(i, 0) <> "" THEN
				checkbox 10, (65 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, count_checkbox)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox
				ref_numb = left(all_clients_array(i, 0),2)
				If InStr(all_membs_with_hcre, ref_numb) <> 0 Then checkbox 200, (65 + (i * 15)), 160, 10, all_clients_array(i, 0), all_clients_array(i, process_checkbox)  'Ignores and blank scanned in persons/strings to avoid a blank checkbox

			End If
		NEXT
		ButtonGroup ButtonPressed
		OkButton 245, 65 + ((total_clients+1) * 15), 50, 15
		CancelButton 300, 65 + ((total_clients+1) * 15), 50, 15
	EndDialog

	Dialog Dialog1
	Cancel_without_confirmation
	check_for_maxis(False)
End If

selected_memb = ""
List_of_HH_membs_to_include = " "					'now we are going to create a list of all the reference numbers of the members that were checked
FOR i = 0 to total_clients
	IF all_clients_array(i, 0) <> "" THEN 						'creates the final array to be used by other scripts.
		HH_memb = left(all_clients_array(i, 0), 2)
		IF all_clients_array(i, count_checkbox) = 1 THEN						'if the person/string has been checked on the dialog then the reference number portion (left 2) will be added to new HH_member_array
			List_of_HH_membs_to_include = List_of_HH_membs_to_include & HH_memb & " "
		END IF
		IF all_clients_array(i, process_checkbox) = 1 THEN
			For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)					'setting the defaults for booleans for each member with HC
				If HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) = HH_memb Then
					HEALTH_CARE_MEMBERS(show_hc_detail_const, hc_memb) = True
					HEALTH_CARE_MEMBERS(HC_major_prog_const, hc_memb) = "MA"
					' If HC_form_name = "Health Care Programs Renewal (DHS-3418)" Then HEALTH_CARE_MEMBERS(HC_eval_process_const, hc_memb) = "Recertification"
					If selected_memb = "" Then selected_memb = hc_memb
				End If
			Next
		End If
	END IF
NEXT
List_of_HH_membs_to_include = trim(List_of_HH_membs_to_include)

case_has_retro_request = False
For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)					'setting the defaults for booleans for each member with HC
	If HEALTH_CARE_MEMBERS(show_hc_detail_const, hc_memb) = False Then HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = False
	If HEALTH_CARE_MEMBERS(member_has_retro_request, hc_memb) = True Then case_has_retro_request = True
Next

MAXIS_footer_month = CM_plus_1_mo					'we are reading CM +1 for information for now.
MAXIS_footer_year = CM_plus_1_yr

month_count = 0										'reading information from STAT using the class in a seperate script
ReDim preserve STAT_INFORMATION(month_count)
'this is set up to be able to read multiple months in the future, if we deterimine that multiple months are needed for this script

Set STAT_INFORMATION(month_count) = new stat_detail

STAT_INFORMATION(month_count).footer_month = MAXIS_footer_month			'setting the month and identifying that we are going to look for only selected members
STAT_INFORMATION(month_count).footer_year = MAXIS_footer_year
STAT_INFORMATION(month_count).LIMIT_MEMBS = TRUE
STAT_INFORMATION(month_count).included_members = List_of_HH_membs_to_include

Call STAT_INFORMATION(month_count).gather_stat_info						'this is where we read

'Now we read STAT/BILS
Call navigate_to_MAXIS_screen("STAT", "BILS")
EMReadScreen existance_check, 1, 2, 73
bils_exist = True
If existance_check = "0" Then bils_exist = False
If bils_exist = True then										'if the panel exists, read the details
	bils_row = 6												'start at the first row
	bils_count = 0
	Do
		bil_ref_numb = ""										'blank out the variables that are read using the BILS function
		bil_date = ""
		bil_serv_code = ""
		bil_gross_amt = ""
		bil_payments = ""
		bil_verif_code = ""
		bil_expense_type_code = ""
		bil_old_priority = ""
		bil_dependent_indicator = ""
		bil_serv_info = ""
		bil_verif_info = ""
		bil_expense_type_info = ""

		ReDim Preserve BILS_ARRAY(last_bils_const, bils_count)		'increment the array of bils
		Call read_BILS_line(bils_row)								'read the line using the function

		BILS_ARRAY(bils_ref_numb_const, bils_count) = bil_ref_numb		'setting the defined variables to the array
		BILS_ARRAY(bils_date_const, bils_count) = bil_date
		BILS_ARRAY(bils_service_code_const, bils_count) = bil_serv_code
		BILS_ARRAY(bils_service_info_const, bils_count) = bil_serv_info
		If bil_serv_code = "" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = ""
		If bil_serv_code = "01" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Health Profsnl"
		If bil_serv_code = "03" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Surgery"
		If bil_serv_code = "04" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Chiropractic"
		If bil_serv_code = "05" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Maternity"
		If bil_serv_code = "07" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Hearing"
		If bil_serv_code = "08" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Vision"
		If bil_serv_code = "09" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Hospital"
		If bil_serv_code = "11" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Hospice"
		If bil_serv_code = "13" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "SNF"
		If bil_serv_code = "14" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Dental"
		If bil_serv_code = "15" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Rx Drug/Supply"
		If bil_serv_code = "16" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Home Health"
		If bil_serv_code = "17" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Diagnostic"
		If bil_serv_code = "18" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Mental Health"
		If bil_serv_code = "19" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Rehabilitation"
		If bil_serv_code = "21" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Med Equip/Supply"
		If bil_serv_code = "22" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Medical Trans"
		If bil_serv_code = "24" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Waivered Serv"
		If bil_serv_code = "25" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Medicare Prem"
		If bil_serv_code = "26" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Dntl/Health Prem"
		If bil_serv_code = "27" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Remedial Care"
		If bil_serv_code = "28" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "MCRE Service"
		If bil_serv_code = "30" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Alternative Care"
		If bil_serv_code = "31" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "MCSHN"
		If bil_serv_code = "32" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Ins Ext Prog"
		If bil_serv_code = "34" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "CW-TCM"
		If bil_serv_code = "37" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Pay-In Spdn"
		If bil_serv_code = "42" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Access Svcs"
		If bil_serv_code = "43" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Chemical Dep"
		If bil_serv_code = "44" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Nutritional Svcs"
		If bil_serv_code = "45" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Transplant"
		If bil_serv_code = "46" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Out-Of-Area Svcs"
		If bil_serv_code = "47" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Copay/Deductible"
		If bil_serv_code = "49" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Prvntv Care"
		If bil_serv_code = "99" Then BILS_ARRAY(bils_service_info_short_const, bils_count) = "Other"
		BILS_ARRAY(bils_gross_amt_const, bils_count) = bil_gross_amt
		BILS_ARRAY(bils_third_payments_const, bils_count) = bil_payments
		BILS_ARRAY(bils_verif_code_const, bils_count) = bil_verif_code
		BILS_ARRAY(bils_verif_info_const, bils_count) = bil_verif_info
		BILS_ARRAY(bils_expense_type_code_const, bils_count) = bil_expense_type_code
		BILS_ARRAY(bils_expense_type_info_const, bils_count) = bil_expense_type_info
		BILS_ARRAY(bils_old_priority_const, bils_count) = bil_old_priority
		BILS_ARRAY(bils_depdnt_ind_const, bils_count) = bil_dependent_indicator

		bils_count = bils_count + 1			'incrementing
		bils_row = bils_row + 1
		If bils_row = 18 Then
			PF20
			EMReadScreen end_of_list, 9, 24, 14
			If end_of_list = "LAST PAGE" Then Exit Do
			bils_row = 6
		End If
		EMReadScreen next_bils_ref_numb, 2, bils_row, 26		'determining when to leave the loop
	Loop until next_bils_ref_numb = "__"
End If

'now we read AREP and SWKR
Call access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)
Call access_SWKR_panel(access_type, swkr_name, swkr_addr_street, swkr_addr_city, swkr_addr_state, swkr_addr_zip, swkr_phone_one, swkr_ext_one, notices_to_swkr_yn)

'Reading SPEC/XFER to find if the case is excluded time
Call navigate_to_MAXIS_screen("SPEC", "XFER")
Call write_value_and_transmit("X", 5, 16)
excluded_time_case = False
EMReadScreen excluded_time_yn, 1, 5, 28
EMReadScreen excluded_time_begin_date, 8, 6, 28
EMReadScreen curr_hc_cty_fin_resp, 2, 14, 39
EMReadScreen curr_ma_cty_fin_resp, 2, 15, 39
curr_hc_cty_fin_resp = replace(curr_hc_cty_fin_resp, "_", "")
curr_ma_cty_fin_resp = replace(curr_ma_cty_fin_resp, "_", "")
EMReadScreen curr_servicing_county, 2, 17, 61
If curr_hc_cty_fin_resp <> "" AND curr_hc_cty_fin_resp <> curr_servicing_county Then
	excluded_time_case = True
	county_of_financial_responsibility = curr_hc_cty_fin_resp
End If
If curr_ma_cty_fin_resp <> "" AND curr_ma_cty_fin_resp <> curr_servicing_county Then
	excluded_time_case = True
	county_of_financial_responsibility = curr_ma_cty_fin_resp
End If

'Pulling the review date from STAT/REVW
Call navigate_to_MAXIS_screen("STAT", "REVW")
EMReadScreen revw_mm, 2, 9, 70
EMReadScreen revw_yy, 2, 9, 76
EMReadScreen revw_date, 8, 9, 70
revw_date = replace(revw_date, " ", "/")
If revw_date = "__/__/__" Then
	revw_date = ""
Else
	revw_date = DateAdd("d", 0, revw_date)
	ma_bc_tikl_date = DateAdd("d", -45, revw_date)
End If

'navigating back to STAT/SUMM for the dialog display
Call navigate_to_MAXIS_screen("STAT", "SUMM")

'here we use what we read using the STAT Class to
' 1. Set informaiton to verifs needed
' 2. identify if certain conditions are met
' 3. create a list of information for verifs to be selected
imig_exists = False
income_source_list = "Select or Type Source"
verifs_needed = ""
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
		income_source_list = income_source_list+chr(9)+"JOB - " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
		income_source_list = income_source_list+chr(9)+"JOB - " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
		income_source_list = income_source_list+chr(9)+"JOB - " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
		income_source_list = income_source_list+chr(9)+"JOB - " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
		income_source_list = income_source_list+chr(9)+"JOB - " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = True Then imig_exists = True
Next
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
		If InStr(income_source_list, "Self Employment") = 0 Then income_source_list = income_source_list+chr(9)+"Self Employment"
		If STAT_INFORMATION(month_ind).stat_busi_one_hc_a_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_busi_two_hc_a_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_busi_three_hc_a_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_income_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Self Employment Income of " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
Next
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & "Income for " & "MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " from " & STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb)
	End If
Next
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_acct_one_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_acct_one_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_acct_one_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_acct_one_location(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_acct_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_acct_two_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_acct_two_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_acct_two_location(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_acct_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_acct_three_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_acct_three_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_acct_three_location(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_acct_four_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_acct_four_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_acct_four_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_acct_four_location(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_acct_five_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_acct_five_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_acct_five_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_acct_five_location(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_secu_one_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_secu_one_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_secu_one_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_secu_one_name(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_secu_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_secu_two_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_secu_two_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_secu_two_name(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_secu_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_secu_three_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_secu_three_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_secu_three_name(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_secu_four_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_secu_four_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_secu_four_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_secu_four_name(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_secu_five_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_secu_five_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; " & STAT_INFORMATION(month_ind).stat_secu_five_type_detail(each_memb) & " Account at " & STAT_INFORMATION(month_ind).stat_secu_five_name(each_memb) & " of MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
Next
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_cars_one_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_cars_one_own_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; Ownership of " & STAT_INFORMATION(month_ind).stat_cars_one_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_model(each_memb) & " owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_cars_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_cars_two_own_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; Ownership of " & STAT_INFORMATION(month_ind).stat_cars_two_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_model(each_memb) & " owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_cars_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_cars_three_own_verif_code(each_memb) = "N" Then verifs_needed = verifs_needed & "; Ownership of " & STAT_INFORMATION(month_ind).stat_cars_three_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_model(each_memb) & " owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_rest_one_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_rest_one_ownership_verif_code(each_memb) = "NO" Then verifs_needed = verifs_needed & "; Ownership of Property (" & STAT_INFORMATION(month_ind).stat_rest_one_type_info(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_rest_one_property_status_info(each_memb) & ") owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_rest_two_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_rest_two_ownership_verif_code(each_memb) = "NO" Then verifs_needed = verifs_needed & "; Ownership of Property (" & STAT_INFORMATION(month_ind).stat_rest_two_type_info(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_rest_two_property_status_info(each_memb) & ") owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
	If STAT_INFORMATION(month_ind).stat_rest_three_exists(each_memb) = True Then
		If STAT_INFORMATION(month_ind).stat_rest_three_ownership_verif_code(each_memb) = "NO" Then verifs_needed = verifs_needed & "; Ownership of Property (" & STAT_INFORMATION(month_ind).stat_rest_three_type_info(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_rest_three_property_status_info(each_memb) & ") owned by MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb)
	End If
Next
If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed)-2)		'formatting the verifs_needed

'This array is to hold notes entered in the dialog BUT because we can't use class parameters to fill information in a dialog, we need to connect them to an array (or a variable)
'This is a bit of a workaround
'The array will hold the information
'The index of the array is defined to the connected class parameter - so the class parameter is a number and identified which array position the information is in
Dim EDITBOX_ARRAY()
ReDim EDITBOX_ARRAY(0)
edit_box_counter = 0
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	'for each class parameter that exists, the counter is set to the class notes and the array size is increased.
	If STAT_INFORMATION(month_ind).stat_emma_exists(each_memb) = True Then
		For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
			If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) Then
				HEALTH_CARE_MEMBERS(HC_major_prog_const, hc_memb) = "EMA"
			End If
		Next
		STAT_INFORMATION(month_ind).stat_emma_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_jobs_one_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_jobs_two_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_jobs_three_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_jobs_four_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_jobs_five_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_busi_one_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_busi_two_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_busi_three_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_unea_one_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_unea_two_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_unea_three_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_unea_four_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_unea_five_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_cash_asset_panel_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_asset_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_imig_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
	If STAT_INFORMATION(month_ind).stat_pben_exists(each_memb) = True Then
		STAT_INFORMATION(month_ind).stat_pben_notes(each_memb) = edit_box_counter
		ReDim preserve EDITBOX_ARRAY(edit_box_counter)
		edit_box_counter = edit_box_counter + 1
	End If
Next
'these ones always exist and don't need an if statement
STAT_INFORMATION(month_ind).stat_jobs_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_busi_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_unea_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_acct_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_cars_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_rest_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_expenses_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_acci_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_insa_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_faci_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1
STAT_INFORMATION(month_ind).stat_other_general_notes = edit_box_counter
ReDim preserve EDITBOX_ARRAY(edit_box_counter)
edit_box_counter = edit_box_counter + 1

'Now we make some other lists and defaults for the verifs dialog
employment_source_list = income_source_list
income_source_list = income_source_list+chr(9)+"Child Support"+chr(9)+"Social Security Income"+chr(9)+"Unemployment Income"+chr(9)+"VA Income"+chr(9)+"Pension"
income_verif_time = "[Enter Time Frame]"
bank_verif_time = "[Enter Time Frame]"
processing_an_application = False
processing_a_recert = False

'These are booleans to decide if we can move on in the scirpt
eval_questions_clear = False
show_err_msg_during_movement = True
'These are where we start this information
page_display = show_member_page
month_ind = 0
Do
	Do
		Do
			Do
				Dialog1 = Empty					'blank out the dialog
				call define_main_dialog			'create the dialog image
				err_msg = ""					'blanking out the error messaging

				prev_page = page_display					'setting what was previously happening on the dialog
				previous_button_pressed = ButtonPressed

				dialog Dialog1					'show the actual dialog

				cancel_confirmation				'this cancels the script if the worker presses 'cancel' or 'ESC'
				Call verification_dialog		'show the verification dialog if the verifs button is pressed.
				Call check_for_errors(eval_questions_clear)								'review entered information to see if there are dialog errors
				Call display_errors(err_msg, False, show_err_msg_during_movement)		'show the error if any exist
			Loop until err_msg = ""
			call dialog_movement				'use the buttons to change the main dialog
		Loop until leave_loop = TRUE			'keep going until the we can leave the loop (the 'complete' button is pressed)
		'this is where we make sure the worker is done entering informaiton.
		proceed_confirm = MsgBox("Have you completed the Health Care Evaluation?" & vbCr & vbCr &_
								 "Once you proceed from this point, there is no opportunity to change information that will be entered in CASE/NOTE." & vbCr & vbCr &_
								 "Press 'No' now if you have additional notes to make or information to review/enter. This will bring you back to the main dailogs." & vbCr &_
								 "Press 'Yes' to confinue to the final part of the interivew (forms)." & vbCr &_
								 "Press 'Cancel' to end the script run.", vbYesNoCancel+ vbQuestion, "Confirm Health Care Evaluation?")
		If proceed_confirm = vbCancel then cancel_confirmation
	Loop Until proceed_confirm = vbYes
	Call check_for_password(are_we_passworded_out)			'make sure we are not passworded out
Loop until are_we_passworded_out = FALSE
Call check_for_MAXIS(False)					'Make sure we are in MAXIS

For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
		If HEALTH_CARE_MEMBERS(HC_eval_process_const, selected_memb) = "Application" Then processing_an_application = True
		If HEALTH_CARE_MEMBERS(HC_eval_process_const, selected_memb) = "Recertification" Then processing_a_recert = True
	End If
Next

If client_delay_check = checked then 'UPDATES PND2 FOR CLIENT DELAY IF CHECKED
	call navigate_to_MAXIS_screen("REPT", "PND2")
	EMGetCursor PND2_row, PND2_col
	EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
	If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
	EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
	If PND2_HC_status_check = "P" then
		EMWriteScreen "x", PND2_row, 3
		transmit
		person_delay_row = 7
		Do
			EMReadScreen person_delay_check, 1, person_delay_row, 39
			If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
			person_delay_row = person_delay_row + 2
		Loop until person_delay_check = " " or person_delay_row > 20
		PF3
	End if
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_check = 0
	End if
End if

'if this application was 4/1/23 or after, we need to ask about STANDARD vs PROTECTED Policy
'For cases that are at recert, we do not need to complete this process during the processing as a bulk script will create this CASE/NOTE
If applied_after_03_23 = True and processing_an_application = True and processing_a_recert = False Then
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 476, 285, "Determine Health Care Policy to Apply"
	DropListBox 160, 195, 275, 45, "Select One..."+chr(9)+"Standard Policy - Changes and Reported information can be acted on"+chr(9)+"Protected Policy - Continuous Coverage applies and not negative action can be taken", policy_to_apply
	DropListBox 160, 215, 275, 35, "Applied on or after 4/1/2023 and no Non-Retro coverage existed in 03/2023.", policy_selection_reason
	ButtonGroup ButtonPressed
		OkButton 360, 260, 50, 15
		CancelButton 415, 260, 50, 15
		PushButton 340, 135, 95, 15, "One Source - COVID", one_source_covid_btn
		PushButton 360, 170, 75, 15, "Knowledge Now", qi_kn_btn
	GroupBox 10, 10, 455, 245, "Policy to Apply to HC Case"
	Text 35, 30, 400, 10, "******************************************************************************************************************************************************************************************************************************************************************************************************"
	Text 180, 40, 105, 10, "*  *  *   SELECT POLICY *  *  *"
	Text 150, 55, 200, 10, "IDENTIFY IF STANDARD OR PROTECTED POLICY APPLY"
	Text 35, 70, 400, 10, "******************************************************************************************************************************************************************************************************************************************************************************************************"
	Text 35, 90, 345, 20, "Since 03/2020 health care eligibility has been maintained under Continuous Coverage rules due to the Public Health Emergency (PHE). With the PHE ending, applied policy will need to be determined for each case."
	Text 35, 115, 290, 20, "If anyone on this case has Non-Retro MA coverage in 03/2023, Protected Policy applies until the first renewal after the end of the PHE."
	Text 90, 140, 250, 10, "Full details of determining which policy applies can be found on One Source"
	Text 35, 160, 275, 10, "Review the case to determine if Standard or Protected Policy Coverage Apply"
	Text 75, 175, 280, 10, "If you need additional support on making this determination, contact Knowledge Now."
	Text 35, 200, 125, 10, "Select the correct policy that applies:"
	Text 65, 220, 95, 10, "Reason to Apply this Policy:"
	Text 100, 235, 265, 10, "This script will create a CASE/NOTE for any case that has Standard Policy Apply."
	EndDialog

	Do
		Do
			err_msg = ""

			dialog Dialog1
			cancel_confirmation

			If policy_to_apply = "Select One..." Then err_msg = err_msg & vbCr & "* Select which policy applies to the members on this case."
			If ButtonPressed = one_source_covid_btn or ButtonPressed = qi_kn_btn Then err_msg = "LOOP" & err_msg
			If ButtonPressed = one_source_covid_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=covidhome"
			If ButtonPressed = qi_kn_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/Quality-Improvement-(QI)-Team.aspx"
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False

	'creating a case note for cases that have STANDARD Policy apply
	If policy_to_apply = "Standard Policy - Changes and Reported information can be acted on" Then
		end_msg = end_msg & vbCr & vbCr & "STANDARD POLICY Now applies to this case." & vbCr & "A CASE/NOTE has been entered to identify the case uses standard policy."
		Call start_a_blank_CASE_NOTE

		Call write_variable_in_CASE_NOTE("~*~*~ MA STANDARD POLICY APPLIES TO THIS CASE ~*~*~")
		Call write_variable_in_CASE_NOTE(policy_selection_reason)' = "Applied on or after 4/1/2023 and no Non-Retro coverage existed in 03/2023."
		Call write_variable_in_CASE_NOTE("**************************************************************************")
		Call write_variable_in_CASE_NOTE("Any future changes or CICs reported can be acted on,")
		Call write_variable_in_CASE_NOTE("even if they result in negative action for Health Care eligibility.")
		Call write_variable_in_CASE_NOTE("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
		Call write_variable_in_CASE_NOTE("Continuous Coverage no longer applies to this case.")
		Call write_variable_in_CASE_NOTE("**************************************************************************")
		Call write_variable_in_CASE_NOTE("Details about this determination can be found in")
		Call write_variable_in_CASE_NOTE("        ONESource in the COVID-19 Page.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)

		PF3
	    Call back_to_SELF
	End If
End If

'If there are verifs_needed listed, we are going to create a CASE/NOTE about verifications needed.
If trim(verifs_needed) <> "" Then
	end_msg = end_msg & vbCr & vbCr & "Verifications were indicated during the Health Care Evaluation and a CASE/NOTE with verification details has been created."

    verif_counter = 1
    verifs_needed = trim(verifs_needed)
    If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
    If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
    If InStr(verifs_needed, ";") <> 0 Then
        verifs_array = split(verifs_needed, ";")
    Else
        verifs_array = array(verifs_needed)
    End If

    Call start_a_blank_CASE_NOTE

    Call write_variable_in_CASE_NOTE("VERIFICATIONS REQUESTED")

    Call write_bullet_and_variable_in_CASE_NOTE("Verif request form sent on", verif_req_form_sent_date)

    Call write_variable_in_CASE_NOTE("---")

    Call write_variable_in_CASE_NOTE("List of all verifications requested:")
    For each verif_item in verifs_array
        verif_item = trim(verif_item)
        If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
        verif_counter = verif_counter + 1
        Call write_variable_with_indent_in_CASE_NOTE(verif_item)
    Next
    If verifs_postponed_checkbox = checked Then
        Call write_variable_in_CASE_NOTE("---")
        Call write_variable_in_CASE_NOTE("There may be verifications that are postponed to allow for the approval of Expedited SNAP.")
    End If
    Call write_variable_in_CASE_NOTE("---")
    Call write_variable_in_CASE_NOTE(worker_signature)

    PF3											'save the note
    EMReadScreen top_note_header, 55, 5, 25

    Call back_to_SELF
End If

'enter TIKLs if requested
If TIKL_check = 1 then Call create_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and set another TIKL reminder.", 45, form_date, False, TIKL_note_text)
If MA_BC_end_of_cert_TIKL_check = checked Then Call create_TIKL("MA-BC recertification is scheduled for " & revw_mm & "/" & revw_yy & ", recertification paperwork needs to be sent manually for this case.", 0, ma_bc_tikl_date, False, MA_BC_TIKL_note_text)

'Here we are creating some variables for the CASE/NOTE
hc_case_determination_status = ""
For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
		curr_memb_status = HEALTH_CARE_MEMBERS(hc_eval_status, the_memb)
		If left(curr_memb_status, 10) = "Incomplete" Then curr_memb_status = "Incomplete"
		If hc_case_determination_status = "" Then hc_case_determination_status = curr_memb_status

		If curr_memb_status = "More Processing Needed" Then hc_case_determination_status = curr_memb_status
		If curr_memb_status = "Incomplete" AND (hc_case_determination_status = "Complete" OR hc_case_determination_status = "Appears Ineligible") Then hc_case_determination_status = curr_memb_status
		If curr_memb_status = "Complete" AND hc_case_determination_status = "Appears Ineligible" Then hc_case_determination_status = curr_memb_status
	End If
Next

'creating a name that is easier to read
If HC_form_name = "Health Care Programs Application for Certain Populations (DHS-3876)" Then short_form = "HC Certain Pops App"
If HC_form_name = "MNsure Application for Health Coverage and Help paying Costs (DHS-6696)" Then short_form = "MNSure HC App"
If HC_form_name = "Health Care Programs Renewal (DHS-3418)" Then short_form = "HC Renewal"
If HC_form_name = "Combined Annual Renewal for Certain Populations (DHS-3727)" Then short_form = "Combined AR"
If HC_form_name = "Application for Payment of Long-Term Care Services (DHS-3531)" Then short_form = "LTC HC App"
If HC_form_name = "Renewal for People Receiving Medical Assistance for Long-Term Care Services (DHS-2128)" Then short_form = "LTC Renewal"
If HC_form_name = "Breast and Cervical Cancer Coverage Group (DHS-3525)" Then short_form = "B/C Cancer App"
If HC_form_name = "Minnesota Family Planning Program Application (DHS-4740)" Then short_form = "MN Family Planning App"
If HC_form_name = "Combined Six Month Report (DHS-5576)" Then short_form = "CSR"


'MAIN CASE/NOTE
start_a_blank_CASE_NOTE
CALL write_variable_in_case_note(form_date & " " & short_form & ": " & hc_case_determination_status)
Call write_bullet_and_variable_in_CASE_NOTE("Form Recvd", HC_form_name)
If ltc_waiver_request_yn = "Yes" Then Call write_variable_in_CASE_NOTE("* This application can be used to request LTC/Waiver services.")
Call write_bullet_and_variable_in_CASE_NOTE("Date Recvd", form_date)
If policy_to_apply = "Protected Policy - Continuous Coverage applies and not negative action can be taken" Then Call write_variable_in_CASE_NOTE("* PROTECTED POLICY from Public Health Emergency still apply at this time.")

If ma_bc_authorization_form_missing_checkbox = unchecked Then
	Call write_bullet_and_variable_in_CASE_NOTE("MA-BC Form Recvd", ma_bc_authorization_form)
	Call write_bullet_and_variable_in_CASE_NOTE("MA-BC Form Date Recvd", ma_bc_authorization_form_date)
End If
Call write_bullet_and_variable_in_CASE_NOTE("Notes", case_details_notes)
If trim(ltc_elig_notes) <> "" or trim(ltc_info_still_needed) <> "" Then
	Call write_variable_in_CASE_NOTE("========================= LTC INFORMATION ==========================")
	Call write_bullet_and_variable_in_CASE_NOTE("ELIG Notes", ltc_elig_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Info Still Needed", ltc_info_still_needed)
End If
Call write_variable_in_CASE_NOTE("========================== PERSON DETAILS ==========================")
For the_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
	If HEALTH_CARE_MEMBERS(show_hc_detail_const, the_memb) = True Then
		Call write_variable_in_CASE_NOTE("MEMB " & HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) & " - " & HEALTH_CARE_MEMBERS(full_name_const, the_memb) & " - Processing: " & HEALTH_CARE_MEMBERS(HC_eval_process_const, the_memb))
		Call write_variable_in_CASE_NOTE("     Status: " & HEALTH_CARE_MEMBERS(hc_eval_status, the_memb))
		If trim(HEALTH_CARE_MEMBERS(hc_eval_notes, the_memb)) <> "" Then Call write_variable_in_CASE_NOTE("     Notes: " & HEALTH_CARE_MEMBERS(hc_eval_notes, the_memb))
		If HEALTH_CARE_MEMBERS(HC_major_prog_const, the_memb) = "None" Then
			Call write_variable_in_CASE_NOTE("     No Health Care Program.")
		Else
			Call write_variable_in_CASE_NOTE("     " & HEALTH_CARE_MEMBERS(HC_major_prog_const, the_memb) & " Basis: " & HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, the_memb))
			If HEALTH_CARE_MEMBERS(HC_basis_of_elig_const, the_memb) = "Breast/Cervical Cancer" Then
				Call write_variable_in_CASE_NOTE("               MA-BC uses Method X Budgeting.")
				Call write_variable_in_CASE_NOTE("               Income/Assets are not counted.")
				If ma_bc_authorization_form_missing_checkbox = checked Then Call write_variable_in_CASE_NOTE("               MA-BC form (" & ma_bc_authorization_form & ") has not been received.")
			End If
		End If
		If trim(HEALTH_CARE_MEMBERS(MA_basis_notes_const, the_memb)) <> "" Then Call write_variable_in_CASE_NOTE("        Notes: " & HEALTH_CARE_MEMBERS(MA_basis_notes_const, the_memb))
		If HEALTH_CARE_MEMBERS(MSP_major_prog_const, the_memb) = "None" Then
			Call write_variable_in_CASE_NOTE("     No Medicare Savings Program.")
		Else
			Call write_variable_in_CASE_NOTE("     MSP Program: " & HEALTH_CARE_MEMBERS(MSP_major_prog_const, the_memb))
			Call write_variable_in_CASE_NOTE("     MSP Basis: " & HEALTH_CARE_MEMBERS(MSP_basis_of_elig_const, the_memb))
		End If
		If trim(HEALTH_CARE_MEMBERS(MSP_basis_notes_const, the_memb)) <> "" Then Call write_variable_in_CASE_NOTE("         Notes: " & HEALTH_CARE_MEMBERS(MSP_basis_notes_const, the_memb))
		If HEALTH_CARE_MEMBERS(member_has_retro_request, the_memb) = True Then
			Call write_variable_in_CASE_NOTE("     RETRO Request back to " & HEALTH_CARE_MEMBERS(hc_cov_date_const, the_memb))
		End If
		'TODO - add MEMB/MEMI information
		For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) Then
				If STAT_INFORMATION(month_ind).stat_imig_exists(each_memb) = True Then
					imig_string = ""
					imig_string = "This resident is a non-citizen; Immigration Status: " & STAT_INFORMATION(month_ind).stat_imig_status_info(each_memb) & ", entry date: " & STAT_INFORMATION(month_ind).stat_imig_entry_date(each_memb) & ", Nationality: " & STAT_INFORMATION(month_ind).stat_imig_nationality_info(each_memb) & "; "
					If STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_code(each_memb) <> "24" AND STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_code(each_memb) <> "__" Then imig_string = imig_string & "LPR Adjusted from " & STAT_INFORMATION(month_ind).stat_imig_LPR_adj_from_info(each_memb) & " on " & STAT_INFORMATION(month_ind).stat_imig_status_verif_code(each_memb) & "; "
					imig_string = imig_string & "Verif: " & STAT_INFORMATION(month_ind).stat_imig_status_verif_info(each_memb) & "; "
					If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_imig_notes(each_memb))) <> "" Then imig_string = imig_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_imig_notes(each_memb))

					Call write_header_and_detail_in_CASE_NOTE("Immigration", imig_string)
				End If
			End If
		Next


		If HEALTH_CARE_MEMBERS(DISA_exists_const, the_memb) = True Then
			disa_string = ""
			disa_string = "HC DISA status: " & HEALTH_CARE_MEMBERS(DISA_hc_status_info_const, the_memb)
			disa_string = disa_string & ", DISA Start Date: " & HEALTH_CARE_MEMBERS(DISA_start_date_const, the_memb)
			If HEALTH_CARE_MEMBERS(DISA_cert_start_const, the_memb) <> "" Then disa_string = disa_string & ", Cert Date Start Date: " & HEALTH_CARE_MEMBERS(DISA_cert_start_const, the_memb)
			disa_string = disa_string & "; "
			If HEALTH_CARE_MEMBERS(DISA_end_date_const, the_memb) <> "" Then disa_string = disa_string & "Disability End Date:: " & HEALTH_CARE_MEMBERS(DISA_end_date_const, the_memb)
			If HEALTH_CARE_MEMBERS(DISA_cert_end_const, the_memb) <> "" Then
				If right(disa_string, 2) = "; " Then disa_string = disa_string & "Cert Date End Date: " & HEALTH_CARE_MEMBERS(DISA_cert_end_const, the_memb)
				If right(disa_string, 2) <> "; " Then disa_string = disa_string & ", Cert Date End Date: " & HEALTH_CARE_MEMBERS(DISA_cert_end_const, the_memb)
			End If
			disa_string = disa_string & "Verif: " & HEALTH_CARE_MEMBERS(DISA_hc_verif_info_const, the_memb) & "; "
			If trim(HEALTH_CARE_MEMBERS(DISA_notes_const, the_memb)) <> "" Then disa_string = disa_string & "Notes: " & HEALTH_CARE_MEMBERS(DISA_notes_const, the_memb) & "; "
			If right(disa_string, 2) <> "; " Then disa_string = disa_string & "; "
			Call write_header_and_detail_in_CASE_NOTE("Disability", disa_string)

			waiver_string = ""
			If HEALTH_CARE_MEMBERS(DISA_waiver_info_const, the_memb) <> "" Then waiver_string = waiver_string & "" & HEALTH_CARE_MEMBERS(DISA_waiver_info_const, the_memb) & "; "
			If trim(HEALTH_CARE_MEMBERS(LTC_waiver_notes_const, selected_memb)) <> "" Then waiver_string = waiver_string & "LTC Notes: " & HEALTH_CARE_MEMBERS(LTC_waiver_notes_const, selected_memb)
			If waiver_string <> "" Then
				If right(waiver_string, 2) <> "; " Then waiver_string = waiver_string & "; "
				Call write_header_and_detail_in_CASE_NOTE("Waiver", waiver_string)
			End If
		End If

		For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
			If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, the_memb) Then
				If STAT_INFORMATION(month_ind).stat_emma_exists(each_memb) = True Then
					emma_string = ""
					emma_string = emma_string & STAT_INFORMATION(month_ind).stat_emma_med_emer_info(each_memb) & "; "
					emma_string = emma_string & "Health Consequence: " & STAT_INFORMATION(month_ind).stat_emma_health_cons_info(each_memb) & "; "
					emma_string = emma_string & "Verif: " & STAT_INFORMATION(month_ind).stat_emma_verif_info(each_memb) & "; "
					emma_string = emma_string & "Begin Date: " & STAT_INFORMATION(month_ind).stat_emma_begin_date(each_memb)
					If STAT_INFORMATION(month_ind).stat_emma_end_date(each_memb) <> "" Then emma_string = emma_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_emma_end_date(each_memb)
					emma_string = emma_string & "; "
					If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_emma_notes(each_memb))) <> "" Then emma_string = emma_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_emma_notes(each_memb))
					Call write_header_and_detail_in_CASE_NOTE("Medical Emergency", emma_string)
				End If
			End If
		Next

		If HEALTH_CARE_MEMBERS(PREG_exists_const, the_memb) = True Then
			preg_string = ""
			preg_string = "Due Date: " & HEALTH_CARE_MEMBERS(PREG_due_date_const, the_memb) & ", Verif:" &  HEALTH_CARE_MEMBERS(PREG_due_date_verif_const, the_memb)
			If HEALTH_CARE_MEMBERS(PREG_multiple_const, the_memb) <> "" Then preg_string = preg_string & ", Multiples: " & HEALTH_CARE_MEMBERS(PREG_multiple_const, the_memb)
			preg_string = preg_string & "; "
			If HEALTH_CARE_MEMBERS(PREG_end_date_const, the_memb) <> "" Then preg_string = preg_string & "Pregnancy End Date: " & HEALTH_CARE_MEMBERS(PREG_end_date_const, the_memb) & ", Verif:" &  HEALTH_CARE_MEMBERS(PREG_end_date_verif_const, the_memb)
			If trim(HEALTH_CARE_MEMBERS(PREG_notes_const, the_memb)) <> "" Then preg_string = preg_string & "Notes: " & HEALTH_CARE_MEMBERS(PREG_notes_const, the_memb)
			Call write_header_and_detail_in_CASE_NOTE("Pregnancy", preg_string)
		End If

		If HEALTH_CARE_MEMBERS(PARE_exists_const, the_memb) = True Then
			pare_string = ""
			pare_string = "Listed as a parent of:" & HEALTH_CARE_MEMBERS(PARE_list_of_children_const, the_memb)
			If trim(HEALTH_CARE_MEMBERS(PARE_notes_const, the_memb)) <> "" Then pare_string = pare_string & "; Notes: " & HEALTH_CARE_MEMBERS(PARE_notes_const, the_memb)
			Call write_header_and_detail_in_CASE_NOTE("Parent", pare_string)
		End If

		If HEALTH_CARE_MEMBERS(MEDI_exists_const, the_memb) = True Then
			medi_string = ""
			If HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, the_memb) <> "" Then
				medi_string = medi_string & "Part A Premium: $ " & HEALTH_CARE_MEMBERS(MEDI_part_A_premium_const, the_memb) & ", Start Date: " & HEALTH_CARE_MEMBERS(MEDI_part_A_start_const, the_memb) & "; "
				medi_string = medi_string & "Part A End Date: " & HEALTH_CARE_MEMBERS(MEDI_part_A_end_const, the_memb) & "; "
			End If
			If HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, the_memb) <> "" Then
				medi_string = medi_string & "Part B Premium: $ " & HEALTH_CARE_MEMBERS(MEDI_part_B_premium_const, the_memb) & ", Start Date: " & HEALTH_CARE_MEMBERS(MEDI_part_B_start_const, the_memb) & "; "
				medi_string = medi_string & "Part B End Date: " & HEALTH_CARE_MEMBERS(MEDI_part_B_end_const, the_memb) & "; "
			End If

			If trim(HEALTH_CARE_MEMBERS(MEDI_notes_const, the_memb)) <> "" Then medi_string = medi_string & "Notes: " & HEALTH_CARE_MEMBERS(MEDI_notes_const, the_memb)
			Call write_header_and_detail_in_CASE_NOTE("Medicare", medi_string)
		Else
			If HEALTH_CARE_MEMBERS(MEDI_application_requred_checkbox_const, the_memb) = checked Then
				If HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) <> "" Then Call write_header_and_detail_in_CASE_NOTE("Medicare", "Application for Medicare is required, referral date: " & HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) & ".")
				If HEALTH_CARE_MEMBERS(MEDI_referral_date_const, each_hh_memb) = "" Then Call write_header_and_detail_in_CASE_NOTE("Medicare", "Application for Medicare is required.")
			End If
		End If
	End If
Next

Call write_variable_in_CASE_NOTE("============================== INCOME ==============================")
income_detail_entered = False
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_jobs_one_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " job at " & STAT_INFORMATION(month_ind).stat_jobs_one_employer_name(each_memb))
		jobs_string = ""
		jobs_string = jobs_string & "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_one_health_care_income_pay_day(each_memb) & " per pay date"
		jobs_string = jobs_string & ", Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_one_main_pay_freq(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_jobs_one_inc_start_date(each_memb) <> "__ __ __" Then jobs_string = jobs_string & "Start date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_one_inc_end_date(each_memb) <> "" Then jobs_string = jobs_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_jobs_one_inc_end_date(each_memb)
		jobs_string = jobs_string & "; "

		If STAT_INFORMATION(month_ind).stat_jobs_one_verif_code(each_memb) = "N" Then
			jobs_string = jobs_string & "No Verification Received; "
		Else
			jobs_string = jobs_string & "Verif: " & STAT_INFORMATION(month_ind).stat_jobs_one_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_one_notes(each_memb))) <> "" Then jobs_string = jobs_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_one_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Job Detail", jobs_string)
	End If

	If STAT_INFORMATION(month_ind).stat_jobs_two_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " job at " & STAT_INFORMATION(month_ind).stat_jobs_two_employer_name(each_memb))
		jobs_string = ""
		jobs_string = jobs_string & "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_two_health_care_income_pay_day(each_memb) & " per pay date"
		jobs_string = jobs_string & ", Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_two_main_pay_freq(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_jobs_two_inc_start_date(each_memb) <> "__ __ __" Then jobs_string = jobs_string & "Start date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_two_inc_end_date(each_memb) <> "" Then jobs_string = jobs_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_jobs_two_inc_end_date(each_memb)
		jobs_string = jobs_string & "; "

		If STAT_INFORMATION(month_ind).stat_jobs_two_verif_code(each_memb) = "N" Then
			jobs_string = jobs_string & "No Verification Received; "
		Else
			jobs_string = jobs_string & "Verif: " & STAT_INFORMATION(month_ind).stat_jobs_two_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_two_notes(each_memb))) <> "" Then jobs_string = jobs_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_two_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Job Detail", jobs_string)
	End If

	If STAT_INFORMATION(month_ind).stat_jobs_three_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " job at " & STAT_INFORMATION(month_ind).stat_jobs_three_employer_name(each_memb))
		jobs_string = ""
		jobs_string = jobs_string & "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_three_health_care_income_pay_day(each_memb) & " per pay date"
		jobs_string = jobs_string & ", Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_three_main_pay_freq(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_jobs_three_inc_start_date(each_memb) <> "__ __ __" Then jobs_string = jobs_string & "Start date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_three_inc_end_date(each_memb) <> "" Then jobs_string = jobs_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_jobs_three_inc_end_date(each_memb)
		jobs_string = jobs_string & "; "

		If STAT_INFORMATION(month_ind).stat_jobs_three_verif_code(each_memb) = "N" Then
			jobs_string = jobs_string & "No Verification Received; "
		Else
			jobs_string = jobs_string & "Verif: " & STAT_INFORMATION(month_ind).stat_jobs_three_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_three_notes(each_memb))) <> "" Then jobs_string = jobs_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_three_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Job Detail", jobs_string)
	End If

	If STAT_INFORMATION(month_ind).stat_jobs_four_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " job at " & STAT_INFORMATION(month_ind).stat_jobs_four_employer_name(each_memb))
		jobs_string = ""
		jobs_string = jobs_string & "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_four_health_care_income_pay_day(each_memb) & " per pay date"
		jobs_string = jobs_string & ", Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_four_main_pay_freq(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_jobs_four_inc_start_date(each_memb) <> "__ __ __" Then jobs_string = jobs_string & "Start date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_four_inc_end_date(each_memb) <> "" Then jobs_string = jobs_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_jobs_four_inc_end_date(each_memb)
		jobs_string = jobs_string & "; "

		If STAT_INFORMATION(month_ind).stat_jobs_four_verif_code(each_memb) = "N" Then
			jobs_string = jobs_string & "No Verification Received; "
		Else
			jobs_string = jobs_string & "Verif: " & STAT_INFORMATION(month_ind).stat_jobs_four_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_four_notes(each_memb))) <> "" Then jobs_string = jobs_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_four_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Job Detail", jobs_string)
	End If

	If STAT_INFORMATION(month_ind).stat_jobs_five_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " job at " & STAT_INFORMATION(month_ind).stat_jobs_five_employer_name(each_memb))
		jobs_string = ""
		jobs_string = jobs_string & "Pay Amount: $ " & STAT_INFORMATION(month_ind).stat_jobs_five_health_care_income_pay_day(each_memb) & " per pay date"
		jobs_string = jobs_string & ", Pay Frequency: " & STAT_INFORMATION(month_ind).stat_jobs_five_main_pay_freq(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_jobs_five_inc_start_date(each_memb) <> "__ __ __" Then jobs_string = jobs_string & "Start date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_jobs_five_inc_end_date(each_memb) <> "" Then jobs_string = jobs_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_jobs_five_inc_end_date(each_memb)
		jobs_string = jobs_string & "; "

		If STAT_INFORMATION(month_ind).stat_jobs_five_verif_code(each_memb) = "N" Then
			jobs_string = jobs_string & "No Verification Received; "
		Else
			jobs_string = jobs_string & "Verif: " & STAT_INFORMATION(month_ind).stat_jobs_five_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_five_notes(each_memb))) <> "" Then jobs_string = jobs_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_five_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Job Detail", jobs_string)
	End If

Next
Call write_bullet_and_variable_in_CASE_NOTE("Job Info", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_jobs_general_notes))

For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_busi_one_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_one_type_info(each_memb))
		busi_string = ""

		If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb) <> "" Then
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			If STAT_INFORMATION(month_ind).stat_busi_one_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_net_inc(each_memb) Then
				busi_string = busi_string & "HC Calculation Method: A and B; "
			Else
				busi_string = busi_string & "HC Calculation Method: B; "
			End If
			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_one_hc_b_income_verif_info(each_memb) & "; "
		Else
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			busi_string = busi_string & "HC Calculation Method: A; "

			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_one_hc_a_income_verif_info(each_memb) & "; "
		End if

		busi_string = busi_string & "Start date: " & STAT_INFORMATION(month_ind).stat_busi_one_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_one_inc_end_date(each_memb) <> "" Then busi_string = busi_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_busi_one_inc_end_date(each_memb)
		busi_string = busi_string & "; "

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_one_notes(each_memb))) <> "" Then busi_string = busi_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_one_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Self Emp Detail", busi_string)

	End If
	If STAT_INFORMATION(month_ind).stat_busi_two_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_two_type_info(each_memb))
		busi_string = ""

		If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb) <> "" Then
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			If STAT_INFORMATION(month_ind).stat_busi_two_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_net_inc(each_memb) Then
				busi_string = busi_string & "HC Calculation Method: A and B; "
			Else
				busi_string = busi_string & "HC Calculation Method: B; "
			End If
			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_two_hc_b_income_verif_info(each_memb) & "; "
		Else
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			busi_string = busi_string & "HC Calculation Method: A; "

			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_two_hc_a_income_verif_info(each_memb) & "; "
		End if

		busi_string = busi_string & "Start date: " & STAT_INFORMATION(month_ind).stat_busi_two_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_two_inc_end_date(each_memb) <> "" Then busi_string = busi_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_busi_two_inc_end_date(each_memb)
		busi_string = busi_string & "; "

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_two_notes(each_memb))) <> "" Then busi_string = busi_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_two_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Self Emp Detail", busi_string)

	End If
		If STAT_INFORMATION(month_ind).stat_busi_three_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " Self Employment Income Type: " & STAT_INFORMATION(month_ind).stat_busi_three_type_info(each_memb))
		busi_string = ""

		If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb) <> "" Then
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			If STAT_INFORMATION(month_ind).stat_busi_three_hc_b_prosp_net_inc(each_memb) = STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_net_inc(each_memb) Then
				busi_string = busi_string & "HC Calculation Method: A and B; "
			Else
				busi_string = busi_string & "HC Calculation Method: B; "
			End If
			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_three_hc_b_income_verif_info(each_memb) & "; "
		Else
			busi_string = busi_string & "Monthly Income: Net $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_net_inc(each_memb)
			busi_string = busi_string & "(Gross: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_gross_inc(each_memb)
			If STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_expenses(each_memb) <> "" Then busi_string = busi_string & " - Expenses: $ " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_prosp_expenses(each_memb)
			busi_string = busi_string & "); "
			busi_string = busi_string & "HC Calculation Method: A; "

			busi_string = busi_string & "Verif: " & STAT_INFORMATION(month_ind).stat_busi_three_hc_a_income_verif_info(each_memb) & "; "
		End if

		busi_string = busi_string & "Start date: " & STAT_INFORMATION(month_ind).stat_busi_three_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_busi_three_inc_end_date(each_memb) <> "" Then busi_string = busi_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_busi_three_inc_end_date(each_memb)
		busi_string = busi_string & "; "

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_three_notes(each_memb))) <> "" Then busi_string = busi_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_three_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Self Emp Detail", busi_string)

	End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Selt Emp Info", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_busi_general_notes))


For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_unea_one_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_one_type_info(each_memb))
		unea_string = ""
		unea_string = unea_string & "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_one_prosp_monthly_gross_income(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_one_inc_start_date(each_memb) <> "__/__/__" Then unea_string = unea_string & ", Start date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb) <> "" Then unea_string = unea_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_unea_one_inc_end_date(each_memb)
		unea_string = unea_string & "; "
		If STAT_INFORMATION(month_ind).stat_unea_one_verif_code(each_memb) = "N" Then
			unea_string = unea_string & "No Verification Received; "
		Else
			unea_string = unea_string & "Verif: " & STAT_INFORMATION(month_ind).stat_unea_one_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_one_notes(each_memb))) <> "" Then unea_string = unea_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_one_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Unearned Detail", unea_string)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_two_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_two_type_info(each_memb))
		unea_string = ""
		unea_string = unea_string & "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_two_prosp_monthly_gross_income(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_two_inc_start_date(each_memb) <> "__/__/__" Then unea_string = unea_string & ", Start date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb) <> "" Then unea_string = unea_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_unea_two_inc_end_date(each_memb)
		unea_string = unea_string & "; "
		If STAT_INFORMATION(month_ind).stat_unea_two_verif_code(each_memb) = "N" Then
			unea_string = unea_string & "No Verification Received; "
		Else
			unea_string = unea_string & "Verif: " & STAT_INFORMATION(month_ind).stat_unea_two_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_two_notes(each_memb))) <> "" Then unea_string = unea_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_two_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Unearned Detail", unea_string)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_three_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_three_type_info(each_memb))
		unea_string = ""
		unea_string = unea_string & "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_three_prosp_monthly_gross_income(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_three_inc_start_date(each_memb) <> "__/__/__" Then unea_string = unea_string & ", Start date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb) <> "" Then unea_string = unea_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_unea_three_inc_end_date(each_memb)
		unea_string = unea_string & "; "
		If STAT_INFORMATION(month_ind).stat_unea_three_verif_code(each_memb) = "N" Then
			unea_string = unea_string & "No Verification Received; "
		Else
			unea_string = unea_string & "Verif: " & STAT_INFORMATION(month_ind).stat_unea_three_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_three_notes(each_memb))) <> "" Then unea_string = unea_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_three_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Unearned Detail", unea_string)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_four_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_four_type_info(each_memb))
		unea_string = ""
		unea_string = unea_string & "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_four_prosp_monthly_gross_income(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_four_inc_start_date(each_memb) <> "__/__/__" Then unea_string = unea_string & ", Start date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb) <> "" Then unea_string = unea_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_unea_four_inc_end_date(each_memb)
		unea_string = unea_string & "; "
		If STAT_INFORMATION(month_ind).stat_unea_four_verif_code(each_memb) = "N" Then
			unea_string = unea_string & "No Verification Received; "
		Else
			unea_string = unea_string & "Verif: " & STAT_INFORMATION(month_ind).stat_unea_four_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_four_notes(each_memb))) <> "" Then unea_string = unea_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_four_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Unearned Detail", unea_string)
	End If
	If STAT_INFORMATION(month_ind).stat_unea_five_exists(each_memb) = True Then
		income_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " income from " & STAT_INFORMATION(month_ind).stat_unea_five_type_info(each_memb))
		unea_string = ""
		unea_string = unea_string & "Monthly Income: $ " & STAT_INFORMATION(month_ind).stat_unea_five_prosp_monthly_gross_income(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_five_inc_start_date(each_memb) <> "__/__/__" Then unea_string = unea_string & ", Start date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_start_date(each_memb)
		If STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb) <> "" Then unea_string = unea_string & ", End Date: " & STAT_INFORMATION(month_ind).stat_unea_five_inc_end_date(each_memb)
		unea_string = unea_string & "; "
		If STAT_INFORMATION(month_ind).stat_unea_five_verif_code(each_memb) = "N" Then
			unea_string = unea_string & "No Verification Received; "
		Else
			unea_string = unea_string & "Verif: " & STAT_INFORMATION(month_ind).stat_unea_five_verif_info(each_memb) & "; "
		End If

		If trim(EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_five_notes(each_memb))) <> "" Then unea_string = unea_string & "Notes: " & EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_five_notes(each_memb))
		Call write_header_and_detail_in_CASE_NOTE("Unearned Detail", unea_string)
	End If


Next
Call write_bullet_and_variable_in_CASE_NOTE("Unearned Info", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_unea_general_notes))
Call write_bullet_and_variable_in_CASE_NOTE("RETRO Income Notes", retro_income_detail)
If income_detail_entered = False Then Call write_variable_in_CASE_NOTE("* No Income for this Case.")


For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_pben_exists(each_memb) = True Then
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " Potential Benefits")
		If STAT_INFORMATION(month_ind).stat_pben_type_code_one(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_one(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_one(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_one(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_one(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_one(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_one(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_one(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_one(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_one(each_memb) & "; " & date_detail)
		End If
		If STAT_INFORMATION(month_ind).stat_pben_type_code_two(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_two(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_two(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_two(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_two(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_two(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_two(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_two(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_two(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_two(each_memb) & "; " & date_detail)
		End If
		If STAT_INFORMATION(month_ind).stat_pben_type_code_three(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_three(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_three(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_three(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_three(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_three(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_three(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_three(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_three(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_three(each_memb) & "; " & date_detail)
		End If
		If STAT_INFORMATION(month_ind).stat_pben_type_code_four(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_four(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_four(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_four(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_four(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_four(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_four(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_four(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_four(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_four(each_memb) & "; " & date_detail)
		End If
		If STAT_INFORMATION(month_ind).stat_pben_type_code_five(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_five(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_five(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_five(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_five(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_five(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_five(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_five(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_five(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_five(each_memb) & "; " & date_detail)
		End If
		If STAT_INFORMATION(month_ind).stat_pben_type_code_six(each_memb) <> "" Then
			date_detail = ""
			If STAT_INFORMATION(month_ind).stat_pben_referral_date_six(each_memb) <> "" Then date_detail = date_detail & "Referral Date: " & STAT_INFORMATION(month_ind).stat_pben_referral_date_six(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_date_applied_six(each_memb) <> "" Then date_detail = date_detail & "Date Applied: " & STAT_INFORMATION(month_ind).stat_pben_date_applied_six(each_memb) & ", "
			If STAT_INFORMATION(month_ind).stat_pben_iaa_date_six(each_memb) <> "" Then date_detail = date_detail & "IAA Date: " & STAT_INFORMATION(month_ind).stat_pben_iaa_date_six(each_memb) & ", "
			date_detail = left(date_detail, len(date_detail)-2)
			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_pben_type_info_six(each_memb), "Status: " & STAT_INFORMATION(month_ind).stat_pben_disp_info_six(each_memb) & " - Verif: " & STAT_INFORMATION(month_ind).stat_pben_verif_info_six(each_memb) & "; " & date_detail)
		End If
		Call write_header_and_detail_in_CASE_NOTE("PBEN Info", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_pben_notes(each_memb)))

	End If
Next

Call write_variable_in_CASE_NOTE("============================== ASSETS ==============================")
If (avs_form_status <> "Select One..." and avs_form_status <> "") OR trim(avs_form_notes) <> "" OR trim(avs_portal_notes) <> "" Then
	Call write_variable_in_CASE_NOTE("-----------------------------------------------------AVS Information")
	If avs_form_status <> "Select One..." Then Call write_bullet_and_variable_in_CASE_NOTE("AVS Authorization Form", avs_form_status)
	Call write_bullet_and_variable_in_CASE_NOTE("Notes", avs_form_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Actions/Details", avs_portal_notes)
	Call write_variable_in_CASE_NOTE("--------------------------------------------------------------------")
End If
asset_detail_entered = False
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_cash_asset_panel_exists(each_memb) = True Then
		asset_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb))

		If STAT_INFORMATION(month_ind).stat_cash_exists(each_memb) = True Then Call write_header_and_detail_in_CASE_NOTE("Cash", "Amount $ "& STAT_INFORMATION(month_ind).stat_cash_balance(each_memb))

		If STAT_INFORMATION(month_ind).stat_acct_one_exists(each_memb) = True Then
			acct_string = ""
			' acct_string = acct_string &
			If STAT_INFORMATION(month_ind).stat_acct_one_location(each_memb) <> "" Then acct_string = acct_string & "At " & STAT_INFORMATION(month_ind).stat_acct_one_location(each_memb)
			acct_string = acct_string & ", Balance: " & STAT_INFORMATION(month_ind).stat_acct_one_balance(each_memb)
			acct_string = acct_string & " as of " & STAT_INFORMATION(month_ind).stat_acct_one_as_of_date(each_memb)
			acct_string = acct_string & "; "
			acct_string = acct_string & "Verif: " & STAT_INFORMATION(month_ind).stat_acct_one_verif_info(each_memb) & ";"

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_acct_one_type_detail(each_memb) & " Account", acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_acct_two_exists(each_memb) = True Then
			acct_string = ""
			' acct_string = acct_string &
			If STAT_INFORMATION(month_ind).stat_acct_two_location(each_memb) <> "" Then acct_string = acct_string & "At " & STAT_INFORMATION(month_ind).stat_acct_two_location(each_memb)
			acct_string = acct_string & ", alance: " & STAT_INFORMATION(month_ind).stat_acct_two_balance(each_memb)
			acct_string = acct_string & " as of " & STAT_INFORMATION(month_ind).stat_acct_two_as_of_date(each_memb)
			acct_string = acct_string & "; "
			acct_string = acct_string & "Verif: " & STAT_INFORMATION(month_ind).stat_acct_two_verif_info(each_memb) & ";"

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_acct_two_type_detail(each_memb) & " Account", acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_acct_three_exists(each_memb) = True Then
			acct_string = ""
			' acct_string = acct_string &
			If STAT_INFORMATION(month_ind).stat_acct_three_location(each_memb) <> "" Then acct_string = acct_string & "At " & STAT_INFORMATION(month_ind).stat_acct_three_location(each_memb)
			acct_string = acct_string & ", Balance: " & STAT_INFORMATION(month_ind).stat_acct_three_balance(each_memb)
			acct_string = acct_string & " as of " & STAT_INFORMATION(month_ind).stat_acct_three_as_of_date(each_memb)
			acct_string = acct_string & "; "
			acct_string = acct_string & "Verif: " & STAT_INFORMATION(month_ind).stat_acct_three_verif_info(each_memb) & ";"

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_acct_three_type_detail(each_memb) & " Account", acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_acct_four_exists(each_memb) = True Then
			acct_string = ""
			' acct_string = acct_string &
			If STAT_INFORMATION(month_ind).stat_acct_four_location(each_memb) <> "" Then acct_string = acct_string & "At " & STAT_INFORMATION(month_ind).stat_acct_four_location(each_memb)
			acct_string = acct_string & ", Balance: " & STAT_INFORMATION(month_ind).stat_acct_four_balance(each_memb)
			acct_string = acct_string & " as of " & STAT_INFORMATION(month_ind).stat_acct_four_as_of_date(each_memb)
			acct_string = acct_string & "; "
			acct_string = acct_string & "Verif: " & STAT_INFORMATION(month_ind).stat_acct_four_verif_info(each_memb) & ";"

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_acct_four_type_detail(each_memb) & " Account", acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_acct_five_exists(each_memb) = True Then
			acct_string = ""
			' acct_string = acct_string &
			If STAT_INFORMATION(month_ind).stat_acct_five_location(each_memb) <> "" Then acct_string = acct_string & "At " & STAT_INFORMATION(month_ind).stat_acct_five_location(each_memb)
			acct_string = acct_string & ", Balance: " & STAT_INFORMATION(month_ind).stat_acct_five_balance(each_memb)
			acct_string = acct_string & " as of " & STAT_INFORMATION(month_ind).stat_acct_five_as_of_date(each_memb)
			acct_string = acct_string & "; "
			acct_string = acct_string & "Verif: " & STAT_INFORMATION(month_ind).stat_acct_five_verif_info(each_memb) & ";"

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_acct_five_type_detail(each_memb) & " Account", acct_string)
		End If

		If STAT_INFORMATION(month_ind).stat_secu_one_exists(each_memb) = True Then
			secu_string = ""
			secu_string = secu_string & "Name: " & STAT_INFORMATION(month_ind).stat_secu_one_name(each_memb) & "; "

			If STAT_INFORMATION(month_ind).stat_secu_one_cash_value(each_memb) <> "" Then secu_string = secu_string & "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_one_cash_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_secu_one_face_value(each_memb) <> "" Then secu_string = secu_string & "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_one_face_value(each_memb)
			secu_string = secu_string & "; "
			If STAT_INFORMATION(month_ind).stat_secu_one_as_of_date(each_memb) <> "__/__/__" Then secu_string = secu_string & "Value as of " & STAT_INFORMATION(month_ind).stat_secu_one_as_of_date(each_memb) & "; "
			secu_string = secu_string & " Verif: " & STAT_INFORMATION(month_ind).stat_secu_one_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_secu_one_type_detail(each_memb), acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_secu_two_exists(each_memb) = True Then
			secu_string = ""
			secu_string = secu_string & "Name: " & STAT_INFORMATION(month_ind).stat_secu_two_name(each_memb) & "; "

			If STAT_INFORMATION(month_ind).stat_secu_two_cash_value(each_memb) <> "" Then secu_string = secu_string & "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_two_cash_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_secu_two_face_value(each_memb) <> "" Then secu_string = secu_string & "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_two_face_value(each_memb)
			secu_string = secu_string & "; "
			If STAT_INFORMATION(month_ind).stat_secu_two_as_of_date(each_memb) <> "__/__/__" Then secu_string = secu_string & "Value as of " & STAT_INFORMATION(month_ind).stat_secu_two_as_of_date(each_memb) & "; "
			secu_string = secu_string & " Verif: " & STAT_INFORMATION(month_ind).stat_secu_two_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_secu_two_type_detail(each_memb), acct_string)
		End If
		If STAT_INFORMATION(month_ind).stat_secu_three_exists(each_memb) = True Then
			secu_string = ""
			secu_string = secu_string & "Name: " & STAT_INFORMATION(month_ind).stat_secu_three_name(each_memb) & "; "

			If STAT_INFORMATION(month_ind).stat_secu_three_cash_value(each_memb) <> "" Then secu_string = secu_string & "Cash (CSV) Value: $ " & STAT_INFORMATION(month_ind).stat_secu_three_cash_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_secu_three_face_value(each_memb) <> "" Then secu_string = secu_string & "Face Value: $ " & STAT_INFORMATION(month_ind).stat_secu_three_face_value(each_memb)
			secu_string = secu_string & "; "
			If STAT_INFORMATION(month_ind).stat_secu_three_as_of_date(each_memb) <> "__/__/__" Then secu_string = secu_string & "Value as of " & STAT_INFORMATION(month_ind).stat_secu_three_as_of_date(each_memb) & "; "
			secu_string = secu_string & " Verif: " & STAT_INFORMATION(month_ind).stat_secu_three_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE(STAT_INFORMATION(month_ind).stat_secu_three_type_detail(each_memb), acct_string)
		End If

		Call write_header_and_detail_in_CASE_NOTE("Liquid Asset Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_asset_notes(each_memb)))
	End If
Next
If asset_detail_entered = False Then Call write_variable_in_CASE_NOTE("* No Liquid Assets for this Case.")
Call write_bullet_and_variable_in_CASE_NOTE("Asset Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_acct_general_notes))

asset_detail_entered = False
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_cars_exists_for_member(each_memb) = True Then
		asset_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb))

		If STAT_INFORMATION(month_ind).stat_cars_one_exists(each_memb) = True Then
			cars_string = ""
			cars_string = cars_string & STAT_INFORMATION(month_ind).stat_cars_one_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_one_model(each_memb)
			cars_string = cars_string & "Use: " & STAT_INFORMATION(month_ind).stat_cars_one_use_info(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_one_hc_clt_benefit_yn(each_memb) <> "" Then cars_string = cars_string & ", HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_one_hc_clt_benefit_yn(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Value: "
			If STAT_INFORMATION(month_ind).stat_cars_one_trade_in_value(each_memb) <> "" Then cars_string = cars_string & "Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_one_trade_in_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_one_loan_value(each_memb) <> "" Then cars_string = cars_string & ", Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_one_loan_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_one_value_source_info(each_memb) <> "" Then cars_string = cars_string & ", Source: " & STAT_INFORMATION(month_ind).stat_cars_one_value_source_info(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Verif: " & STAT_INFORMATION(month_ind).stat_cars_one_own_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE("Vehicle", cars_string)
		End If
		If STAT_INFORMATION(month_ind).stat_cars_two_exists(each_memb) = True Then
			cars_string = ""
			cars_string = cars_string & STAT_INFORMATION(month_ind).stat_cars_two_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_two_model(each_memb)
			cars_string = cars_string & "Use: " & STAT_INFORMATION(month_ind).stat_cars_two_use_info(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_two_hc_clt_benefit_yn(each_memb) <> "" Then cars_string = cars_string & ", HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_two_hc_clt_benefit_yn(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Value: "
			If STAT_INFORMATION(month_ind).stat_cars_two_trade_in_value(each_memb) <> "" Then cars_string = cars_string & "Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_two_trade_in_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_two_loan_value(each_memb) <> "" Then cars_string = cars_string & ", Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_two_loan_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_two_value_source_info(each_memb) <> "" Then cars_string = cars_string & ", Source: " & STAT_INFORMATION(month_ind).stat_cars_two_value_source_info(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Verif: " & STAT_INFORMATION(month_ind).stat_cars_two_own_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE("Vehicle", cars_string)
		End If
		If STAT_INFORMATION(month_ind).stat_cars_three_exists(each_memb) = True Then
			cars_string = ""
			cars_string = cars_string & STAT_INFORMATION(month_ind).stat_cars_three_year(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_make(each_memb) & " " & STAT_INFORMATION(month_ind).stat_cars_three_model(each_memb)
			cars_string = cars_string & "Use: " & STAT_INFORMATION(month_ind).stat_cars_three_use_info(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_three_hc_clt_benefit_yn(each_memb) <> "" Then cars_string = cars_string & ", HC Client Benefit: " & STAT_INFORMATION(month_ind).stat_cars_three_hc_clt_benefit_yn(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Value: "
			If STAT_INFORMATION(month_ind).stat_cars_three_trade_in_value(each_memb) <> "" Then cars_string = cars_string & "Trade In: $ " & STAT_INFORMATION(month_ind).stat_cars_three_trade_in_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_three_loan_value(each_memb) <> "" Then cars_string = cars_string & ", Loan: $ " & STAT_INFORMATION(month_ind).stat_cars_three_loan_value(each_memb)
			If STAT_INFORMATION(month_ind).stat_cars_three_value_source_info(each_memb) <> "" Then cars_string = cars_string & ", Source: " & STAT_INFORMATION(month_ind).stat_cars_three_value_source_info(each_memb)
			cars_string = cars_string & "; "
			cars_string = cars_string & "Verif: " & STAT_INFORMATION(month_ind).stat_cars_three_own_verif_info(each_memb) & "; "

			Call write_header_and_detail_in_CASE_NOTE("Vehicle", cars_string)
		End If

	End if
Next
Call write_bullet_and_variable_in_CASE_NOTE("Vehicle Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_cars_notes))

For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_rest_exists_for_member(each_memb) = True Then
		asset_detail_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb))

		If STAT_INFORMATION(month_ind).stat_rest_one_exists(each_memb) = True Then
			rest_string = ""
			rest_string = rest_string & STAT_INFORMATION(month_ind).stat_rest_one_type_info(each_memb)
			rest_string = rest_string & ", Property Status: " & STAT_INFORMATION(month_ind).stat_rest_one_property_status_info(each_memb) & "; "
			rest_string = rest_string & "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_property_ownership_info(each_memb) & "; "
			If STAT_INFORMATION(month_ind).stat_rest_one_market_value(each_memb) <> "" Then
				rest_string = rest_string & "Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_one_market_value(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_value_verif_info(each_memb) & "; "
			End If
			If STAT_INFORMATION(month_ind).stat_rest_one_amount_owed(each_memb) <> "" Then
				rest_string = rest_string & "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_one_amount_owed(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_one_owed_verif_info(each_memb) & "; "
			End If

			Call write_header_and_detail_in_CASE_NOTE("Real Estate", rest_string)
		End If
		If STAT_INFORMATION(month_ind).stat_rest_two_exists(each_memb) = True Then
			rest_string = ""
			rest_string = rest_string & STAT_INFORMATION(month_ind).stat_rest_two_type_info(each_memb)
			rest_string = rest_string & ", Property Status: " & STAT_INFORMATION(month_ind).stat_rest_two_property_status_info(each_memb) & "; "
			rest_string = rest_string & "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_property_ownership_info(each_memb) & "; "
			If STAT_INFORMATION(month_ind).stat_rest_two_market_value(each_memb) <> "" Then
				rest_string = rest_string & "Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_two_market_value(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_value_verif_info(each_memb) & "; "
			End If
			If STAT_INFORMATION(month_ind).stat_rest_two_amount_owed(each_memb) <> "" Then
				rest_string = rest_string & "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_two_amount_owed(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_two_owed_verif_info(each_memb) & "; "
			End If

			Call write_header_and_detail_in_CASE_NOTE("Real Estate", rest_string)
		End If
		If STAT_INFORMATION(month_ind).stat_rest_three_exists(each_memb) = True Then
			rest_string = ""
			rest_string = rest_string & STAT_INFORMATION(month_ind).stat_rest_three_type_info(each_memb)
			rest_string = rest_string & ", Property Status: " & STAT_INFORMATION(month_ind).stat_rest_three_property_status_info(each_memb) & "; "
			rest_string = rest_string & "Ownership Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_property_ownership_info(each_memb) & "; "
			If STAT_INFORMATION(month_ind).stat_rest_three_market_value(each_memb) <> "" Then
				rest_string = rest_string & "Market Value: $ " & STAT_INFORMATION(month_ind).stat_rest_three_market_value(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_value_verif_info(each_memb) & "; "
			End If
			If STAT_INFORMATION(month_ind).stat_rest_three_amount_owed(each_memb) <> "" Then
				rest_string = rest_string & "Amount Owed: $ " & STAT_INFORMATION(month_ind).stat_rest_three_amount_owed(each_memb)
				rest_string = rest_string & ", Verif: " & STAT_INFORMATION(month_ind).stat_rest_three_owed_verif_info(each_memb) & "; "
			End If

			Call write_header_and_detail_in_CASE_NOTE("Real Estate", rest_string)
		End If

	End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Real Estate Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_rest_notes))
Call write_bullet_and_variable_in_CASE_NOTE("RETRO Asset Notes", retro_asset_detail)
If asset_detail_entered = False Then Call write_variable_in_CASE_NOTE("* No vehicles or real estate for this Case.")

Call write_variable_in_CASE_NOTE("===================== EXPENSES and DEDUCTIONS ======================")
expense_info_entered = False
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_pded_exists(each_memb) = True Then
		expense_info_entered = True
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " - Program Deductions")
		If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) <> "_" Then
			pickle_string = ""
			If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) = "1" then pickle_string = pickle_string & "Eligible: "
			If STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_yn(each_memb) = "2" then pickle_string = pickle_string & "POTENTIALLY Eligible: "
			pickle_string = pickle_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_pickle_disregard_amt(each_memb) & " Disregard Amount; "
			pickle_string = pickle_string & "Current RSDI $ " & STAT_INFORMATION(month_ind).stat_pded_pickle_curr_RSDI(each_memb) & " less Threshold RSDI $ " & STAT_INFORMATION(month_ind).stat_pded_pickle_threshold_RSDI(each_memb) & "; "
			pickle_string = pickle_string & "Based on Threshold Date: " & STAT_INFORMATION(month_ind).stat_pded_pickle_threshold_date(each_memb) & "; "
			Call write_header_and_detail_in_CASE_NOTE("PICKLE Disregard", pickle_string)
		End If

		other_ded_string = ""
		If STAT_INFORMATION(month_ind).stat_pded_disa_widow_deducation_yn(each_memb) = "Y" Then other_ded_string = other_ded_string & "Disabled Widow/ers Deduction Applied; "
		If STAT_INFORMATION(month_ind).stat_pded_disa_adult_child_disregard_yn(each_memb) = "Y" Then other_ded_string = other_ded_string & "Disabled Adult Child Disregard applied; "
		If STAT_INFORMATION(month_ind).stat_pded_widow_deducation_yn(each_memb) = "Y" Then other_ded_string = other_ded_string & "Widow/ers Deduction applied; "
		If STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_amt(each_memb) <> "" Then other_ded_string = other_ded_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_amt(each_memb) & " Unearned Income Deduction Applied, Reason: " & STAT_INFORMATION(month_ind).stat_pded_other_unea_deduction_reason(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_amt(each_memb) <> "" Then other_ded_string = other_ded_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_amt(each_memb) & " Earned Income Deduction Applied, Reason: " & STAT_INFORMATION(month_ind).stat_pded_other_earned_deduction_reason(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_pded_disa_student_child_disregard_yn(each_memb) = "Y" Then other_ded_string = other_ded_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_disa_student_child_disregard_amt(each_memb) & " Blind/Disabled Student Child Disregard; "
		Call write_header_and_detail_in_CASE_NOTE("Other Deductions", other_ded_string)

		If STAT_INFORMATION(month_ind).stat_pded_extend_ma_epd_limits_yn(each_memb) = "Y" Then Call write_variable_in_CASE_NOTE("     MA-EPD Income/Asset Limits Extended")
		If STAT_INFORMATION(month_ind).stat_pded_PASS_begin_date(each_memb) <> "" Then
			pass_string = ""
			pass_string = pass_string & "Begin Date: " & STAT_INFORMATION(month_ind).stat_pded_PASS_begin_date(each_memb)
			If STAT_INFORMATION(month_ind).stat_pded_PASS_end_date(each_memb) <> "" Then pass_string = pass_string & " - End Date: " & STAT_INFORMATION(month_ind).stat_pded_PASS_end_date(each_memb)
			pass_string = pass_string & "; "
			If STAT_INFORMATION(month_ind).stat_pded_PASS_earned_excluded(each_memb) <> "" Then pass_string = pass_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_earned_excluded(each_memb) & " - Earned Income Excluded; "
			If STAT_INFORMATION(month_ind).stat_pded_PASS_unea_excluded(each_memb) <> "" Then pass_string = pass_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_unea_excluded(each_memb) & " - Unearned Income Excluded; "
			If STAT_INFORMATION(month_ind).stat_pded_PASS_assets_excluded(each_memb) <> "" Then pass_string = pass_string & "$ " & STAT_INFORMATION(month_ind).stat_pded_PASS_assets_excluded(each_memb) & " - Assets Excluded; "
			Call write_header_and_detail_in_CASE_NOTE("PASS Plan", pass_string)
		End If
		If STAT_INFORMATION(month_ind).stat_pded_guardianship_fee(each_memb) <> "" Then Call write_variable_in_CASE_NOTE("     $ " & STAT_INFORMATION(month_ind).stat_pded_guardianship_fee(each_memb) & " Guardianship Fee reduced from income.")
		If STAT_INFORMATION(month_ind).stat_pded_rep_payee_fee(each_memb) <> "" Then Call write_variable_in_CASE_NOTE("     $ " & STAT_INFORMATION(month_ind).stat_pded_rep_payee_fee(each_memb) & " Rep Payee Fee reduced from income.")
	End If
Next

For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_coex_exists(each_memb) = True Then
		expense_info_entered = True

		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " - Court Ordered Expenses")
		coex_string = ""
		coex_string = coex_string & "$ " & STAT_INFORMATION(month_ind).stat_coex_total_prosp_amt(each_memb) & " TOTAL Expense; "
		If STAT_INFORMATION(month_ind).stat_coex_support_prosp_amt(each_memb) <> "" Then coex_string = coex_string & "$ " & STAT_INFORMATION(month_ind).stat_coex_support_prosp_amt(each_memb) & " SUPPORT Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_support_verif_info(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_coex_alimony_prosp_amt(each_memb) <> "" Then coex_string = coex_string & "$ " & STAT_INFORMATION(month_ind).stat_coex_alimony_prosp_amt(each_memb) & " ALIMONY Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_alimony_verif_info(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_coex_tax_dep_prosp_amt(each_memb) <> "" Then coex_string = coex_string & "$ " & STAT_INFORMATION(month_ind).stat_coex_tax_dep_prosp_amt(each_memb) & " TAX DEPENDENT Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_tax_dep_verif_info(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_coex_other_prosp_amt(each_memb) <> "" Then coex_string = coex_string & "$ " & STAT_INFORMATION(month_ind).stat_coex_other_prosp_amt(each_memb) & " OTHER Expense - Verif: " & STAT_INFORMATION(month_ind).stat_coex_other_verif_info(each_memb) & "; "
		Call write_header_and_detail_in_CASE_NOTE("COEX Info", coex_string)
	End If
Next

For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_dcex_exists(each_memb) = True Then
		expense_info_entered = True

		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " - Dependent Care Expenses")
		dcex_string = ""
		dcex_string = dcex_string & "Provider: " &  STAT_INFORMATION(month_ind).stat_dcex_provider(each_memb) & " - Reason: " & STAT_INFORMATION(month_ind).stat_dcex_reason_info(each_memb) & "; "
		If InStr(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb), ",") <> 0 Then
			dcex_child_array = split(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb), ",")
			dcex_amount_array = split(STAT_INFORMATION(month_ind).stat_dcex_prosp_amt_list(each_memb), ",")
			dcex_verif_array = split(STAT_INFORMATION(month_ind).stat_dcex_verif_info_list(each_memb), ",")
		Else
			dcex_child_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_child_list(each_memb))
			dcex_amount_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_prosp_amt_list(each_memb))
			dcex_verif_array = ARRAY(STAT_INFORMATION(month_ind).stat_dcex_verif_info_list(each_memb))
		End If
		For dcex_child = 0 to UBound(dcex_child_array)
			dcex_string = dcex_string & "$ " & dcex_amount_array(dcex_child) & " for MEMB " & dcex_child_array(dcex_child) & ", Verif: " & dcex_verif_array(dcex_child) & "; "
		Next
		dcex_child_array = ""
		dcex_amount_array = ""
		dcex_verif_array = ""
		Call write_header_and_detail_in_CASE_NOTE("DCEX Info", dcex_string)
	End If
Next
If expense_info_entered = False Then Call write_variable_in_CASE_NOTE("* No expenses or deductions for this Case.")
Call write_bullet_and_variable_in_CASE_NOTE("RETRO Expense Notes", retro_expense_detail)
Call write_bullet_and_variable_in_CASE_NOTE("Expense Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_expenses_general_notes))

Call write_variable_in_CASE_NOTE("=========================== OTHER INFO =============================")
Call write_bullet_and_variable_in_CASE_NOTE("RETRO Notes", retro_notes)
For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_acci_exists(each_memb) = True Then
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " - Accident Information")
		acci_string = ""
		acci_string = acci_string & "Injury Date " & STAT_INFORMATION(month_ind).stat_acci_injury_date(each_memb) & ". Medical cooperation: " & STAT_INFORMATION(month_ind).stat_acci_med_coop_yn(each_memb) & "; "
		acci_string = acci_string & "Accident Type: " & STAT_INFORMATION(month_ind).stat_acci_type_info(each_memb) & ". Involving MEMBS " & STAT_INFORMATION(month_ind).stat_acci_ref_numbers_list(each_memb)
		Call write_header_and_detail_in_CASE_NOTE("Details", acci_string)
	End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Accident Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_acci_notes))

For each_panel = 0 to UBound(STAT_INFORMATION(month_ind).stat_insa_exists)
	If STAT_INFORMATION(month_ind).stat_insa_exists(each_panel) = True Then
		Call write_variable_in_CASE_NOTE("Other Health Insurance at " & STAT_INFORMATION(month_ind).stat_insa_insurance_co(each_panel) & " coop with OHI: " & STAT_INFORMATION(month_ind).stat_insa_coop_OHI_yn(each_panel) & " CEHI coop: " & STAT_INFORMATION(month_ind).stat_insa_coop_cost_effective_yn(each_panel))
		Call write_header_and_detail_in_CASE_NOTE("Covered MEMBS", STAT_INFORMATION(month_ind).stat_insa_covered_pers_list(each_panel))
		If STAT_INFORMATION(month_ind).stat_insa_good_cause_code(each_panel) <> "_" And STAT_INFORMATION(month_ind).stat_insa_good_cause_code(each_panel) <> "N" Then
			Call write_header_and_detail_in_CASE_NOTE("Good Cause", STAT_INFORMATION(month_ind).stat_insa_good_cause_info(each_panel) & " - Claim Date: " & STAT_INFORMATION(month_ind).stat_insa_good_cause_claim_date(each_panel) & " - Evidence: " & STAT_INFORMATION(month_ind).stat_insa_coop_cost_effective_yn(each_panel))
		End If
	End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Insurance Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_insa_notes))

For each_memb = 0 to UBound(STAT_INFORMATION(month_ind).stat_memb_ref_numb)
	If STAT_INFORMATION(month_ind).stat_faci_exists(each_memb) = True and STAT_INFORMATION(month_ind).stat_faci_currently_in_facility(each_memb) = True Then
		Call write_variable_in_CASE_NOTE("MEMB " & STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) & " - " & STAT_INFORMATION(month_ind).stat_memb_full_name_no_initial(each_memb) & " is in a facility")
		faci_string = ""
		faci_string = faci_string & "Name: " & STAT_INFORMATION(month_ind).stat_faci_name(each_memb)
		faci_string = faci_string & ", Type: " & STAT_INFORMATION(month_ind).stat_faci_type_info(each_memb)
		faci_string = faci_string & ", In Date: " & STAT_INFORMATION(month_ind).stat_faci_date_in(each_memb) & "; "
		If STAT_INFORMATION(month_ind).stat_faci_waiver_type_info(each_memb) <> "" Then
			faci_string = faci_string & "Facility Waiver Type: " & STAT_INFORMATION(month_ind).stat_faci_waiver_type_info(each_memb) & "; "
		End If
		If STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb) <> "" Then
			faci_string = faci_string & "LTC Ineligible Reason: " & STAT_INFORMATION(month_ind).stat_faci_LTC_inelig_reason_info(each_memb) & "; "
		End If
		Call write_header_and_detail_in_CASE_NOTE("Facility Info", faci_string)

		For hc_memb = 0 to UBound(HEALTH_CARE_MEMBERS, 2)
			If STAT_INFORMATION(month_ind).stat_memb_ref_numb(each_memb) = HEALTH_CARE_MEMBERS(ref_numb_const, hc_memb) Then
				Call write_header_and_detail_in_CASE_NOTE("LTC FACI Notes", HEALTH_CARE_MEMBERS(LTC_facility_notes_const, hc_memb))
			End If
		Next

		If excluded_time_case = True Then
			Call write_variable_in_CASE_NOTE("* EXCLUDED TIME CASE - County of Financial Responsibility: " & county_of_financial_responsibility)
		End If
	End If
Next
Call write_bullet_and_variable_in_CASE_NOTE("Facility Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_faci_notes))
Call write_bullet_and_variable_in_CASE_NOTE("Other Notes", EDITBOX_ARRAY(STAT_INFORMATION(month_ind).stat_other_general_notes))

If bils_exist = True Then
	Call write_variable_in_CASE_NOTE("Medical Bill Information exists on case.")
	first_bil = True
	For each_bil = 0 to UBound(BILS_ARRAY, 2)
		If BILS_ARRAY(bils_checkbox, each_bil) = checked Then
			If first_bil = True Then
				Call write_variable_in_CASE_NOTE("  Person  Date     Gross     Service          Type Verif")
				first_bil = False
			End If
			bill_line = ""
			bill_line = "  MEMB " & BILS_ARRAY(bils_ref_numb_const, each_bil)
			bill_line = bill_line & " " & BILS_ARRAY(bils_date_const, each_bil)
			bill_line = bill_line & " $ " & right(space(7)&BILS_ARRAY(bils_gross_amt_const, each_bil), 7)
			bill_line = bill_line & " " & left(BILS_ARRAY(bils_service_info_short_const, each_bil)&space(16), 16)
			bill_line = bill_line & " " & BILS_ARRAY(bils_expense_type_code_const, each_bil)
			bill_line = bill_line & "    " & left(BILS_ARRAY(bils_verif_info_const, each_bil)&space(25), 25)
			Call write_variable_in_CASE_NOTE(bill_line)
		End If
	Next
	Call write_bullet_and_variable_in_CASE_NOTE("Bills Notes",bils_notes)
	If first_bil = False Then call write_variable_in_CASE_NOTE("-----------------------------------------------------------------------------")
End If

If arep_name <> "" Then
	arep_string = arep_name & ", Notices to AREP: " & forms_to_arep & ", MMIS Mail to AREP: " & mmis_mail_to_arep
	Call write_bullet_and_variable_in_CASE_NOTE("Authorized Rep", arep_string)
End If

If swkr_name <> "" Then
	swkr_string = swkr_name & ", Notices to SWKR: " & notices_to_swkr_yn
	Call write_bullet_and_variable_in_CASE_NOTE("Social Worker", swkr_string)
End If

If app_sig_status = "Yes - All required signatures are on the application" Then
	Call write_variable_in_CASE_NOTE("* Application correctly signed and dated.")
Else
	Call write_bullet_and_variable_in_CASE_NOTE("Application was missing", app_sig_notes)
End If
If trim(verifs_needed) <> "" Then Call write_variable_in_CASE_NOTE("** VERIFICATIONS REQUESTED - See previous case note for detail")
If trim(verifs_needed) = "" Then Call write_variable_in_CASE_NOTE("* No Verifications listed for this case.")

IF client_delay_check = checked THEN CALL write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
If TIKL_check = checked then Call write_variable_in_case_note(TIKL_note_text)
If MA_BC_end_of_cert_TIKL_check = checked Then
	MA_BC_TIKL_note_text = replace(MA_BC_TIKL_note_text, ", 0 day return", "")
	Call write_variable_in_case_note(MA_BC_TIKL_note_text & " TIKL to send Recert forms 45 days before REVW.")
End If

' MEMB XX - NAME - Status:
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)

end_msg = "Health Care Evaluation has been completed and entered in CASE/NOTE." & end_msg
Call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/17/2023
'--Tab orders reviewed & confirmed----------------------------------------------04/17/2023
'--Mandatory fields all present & Reviewed--------------------------------------04/17/2023
'--All variables in dialog match mandatory fields-------------------------------04/17/2023
'Review dialog names for content and content fit in dialog----------------------04/17/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/17/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------04/17/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/17/2023
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------04/17/2023
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/17/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------04/17/2023
'--Out-of-County handling reviewed----------------------------------------------04/17/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/17/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------04/17/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/17/2023
'--Incrementors reviewed (if necessary)-----------------------------------------04/17/2023
'--Denomination reviewed -------------------------------------------------------04/17/2023
'--Script name reviewed---------------------------------------------------------04/17/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------04/17/2023
'--comment Code-----------------------------------------------------------------04/17/2023
'--Update Changelog for release/update------------------------------------------04/17/2023
'--Remove testing message boxes-------------------------------------------------04/17/2023
'--Remove testing code/unnecessary code-----------------------------------------04/17/2023
'--Review/update SharePoint instructions----------------------------------------04/17/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/18/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------04/18/2023
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A