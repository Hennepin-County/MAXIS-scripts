'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INTERVIEW SUMMARY.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
' run_locally = TRUE
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
call changelog_update("09/30/2024", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================

' const ref_number					= 0
' const access_denied					= 1
' const full_name_const				= 2
' const last_name_const				= 3
' const first_name_const				= 4
' const mid_initial					= 5
' const other_names					= 6
' const age							= 7
' const date_of_birth					= 8
' const ssn							= 9
' const ssn_verif						= 10
' const birthdate_verif				= 11
' const gender						= 12
' const race							= 13
' const spoken_lang					= 14
' const written_lang					= 15
' const interpreter					= 16
' const alias_yn						= 17
' const ethnicity_yn					= 18
' const id_verif						= 19
' const rel_to_applcnt				= 20
' const cash_minor					= 21
' const snap_minor					= 22
' const marital_status				= 23
' const spouse_ref					= 24
' const spouse_name					= 25
' const last_grade_completed 			= 26
' const citizen						= 27
' const other_st_FS_end_date 			= 28
' const in_mn_12_mo					= 29
' const residence_verif				= 30
' const mn_entry_date					= 31
' const former_state					= 32
' const fs_pwe						= 33
' const button_one					= 34
' const button_two					= 35
' const imig_status 					= 36
' const clt_has_sponsor				= 37
' const client_verification			= 38
' const client_verification_details	= 39
' const client_notes					= 40
' const intend_to_reside_in_mn		= 41
' const race_a_checkbox				= 42
' const race_b_checkbox				= 43
' const race_n_checkbox				= 44
' const race_p_checkbox				= 45
' const race_w_checkbox				= 46
' const snap_req_checkbox				= 47
' const cash_req_checkbox				= 48
' const emer_req_checkbox				= 49
' const none_req_checkbox				= 50
' const ssn_no_space					= 51
' const edrs_msg						= 52
' const edrs_match					= 53
' const edrs_notes 					= 54
' const ignore_person                 = 55
' const pers_in_maxis                 = 56
' const memb_is_caregiver             = 57
' const cash_request_const            = 58
' const hours_per_week_const          = 59
' const exempt_from_ed_const          = 60
' const comply_with_ed_const          = 61
' const orientation_needed_const      = 62
' const orientation_done_const        = 63
' const orientation_exempt_const      = 64
' const exemption_reason_const        = 65
' const emps_exemption_code_const     = 66
' const choice_form_done_const        = 67
' const orientation_notes             = 68
' const last_const					= 69

' Dim HH_MEMB_ARRAY()
' ReDim HH_MEMB_ARRAY(last_const, 0)
'===========================================================================================================================

'FUNCTIONS =================================================================================================================

function access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)

	Call navigate_to_MAXIS_screen("STAT", "AREP")

	EMReadScreen arep_name, 37, 4, 32
	arep_name = replace(arep_name, "_", "")
	If arep_name <> "" Then
		EMReadScreen arep_street_one, 22, 5, 32
		EMReadScreen arep_street_two, 22, 6, 32
		EMReadScreen arep_addr_city, 15, 7, 32
		EMReadScreen arep_addr_state, 2, 7, 55
		EMReadScreen arep_addr_zip, 5, 7, 64

		arep_street_one = replace(arep_street_one, "_", "")
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

function assess_caf_1_expedited_questions(expedited_screening)
	If IsNumeric(exp_q_1_income_this_month) = False Then exp_q_1_income_this_month = 0
	If IsNumeric(exp_q_2_assets_this_month) = False Then exp_q_2_assets_this_month = 0
	If IsNumeric(exp_q_3_rent_this_month) = False Then exp_q_3_rent_this_month = 0

	exp_q_1_income_this_month = FormatNumber(exp_q_1_income_this_month, 2, -1, 0, -1)
	exp_q_2_assets_this_month = FormatNumber(exp_q_2_assets_this_month, 2, -1, 0, -1)
	exp_q_3_rent_this_month = FormatNumber(exp_q_3_rent_this_month, 2, -1, 0, -1)

	exp_q_4_utilities_this_month = 0
	If caf_exp_pay_heat_checkbox = checked OR caf_exp_pay_ac_checkbox = checked Then
		exp_q_4_utilities_this_month = heat_AC_amt
	Else
		If caf_exp_pay_electricity_checkbox = checked Then exp_q_4_utilities_this_month = exp_q_4_utilities_this_month + electric_amt
		If caf_exp_pay_phone_checkbox = checked Then exp_q_4_utilities_this_month = exp_q_4_utilities_this_month + phone_amt
	End If
	exp_q_4_utilities_this_month = FormatNumber(exp_q_4_utilities_this_month, 2, -1, 0, -1)

	caf_1_resources = exp_q_1_income_this_month + exp_q_2_assets_this_month
	caf_1_expenses = exp_q_3_rent_this_month + exp_q_4_utilities_this_month

	expedited_screening = "CAF 1 Information does NOT appear Expedited"

	If exp_q_1_income_this_month < 150 AND exp_q_2_assets_this_month <= 100 Then expedited_screening = "CAF 1 Information APPEARS EXPEDITED"
	If caf_1_resources < caf_1_expenses Then expedited_screening = "CAF 1 Information APPEARS EXPEDITED"

	exp_q_1_income_this_month = exp_q_1_income_this_month & ""
	exp_q_2_assets_this_month = exp_q_2_assets_this_month & ""
	exp_q_3_rent_this_month = exp_q_3_rent_this_month & ""
end function

function check_for_errors(interview_questions_clear)
	who_are_we_completing_the_interview_with = trim(who_are_we_completing_the_interview_with)
	If who_are_we_completing_the_interview_with = "Select or Type" Or who_are_we_completing_the_interview_with = "" Then err_msg = err_msg & "~!~" & "1 ^* Who are you interviewing with?##~##   - Select or enter the name of the person you are completing the interview with.##~##"
	If how_are_we_completing_the_interview = "Select or Type" Or how_are_we_completing_the_interview = "" Then err_msg = err_msg & "~!~" & "1 ^* Interview via##~##   - Select or enter the method the interview is being conducted.##~##"
	If trim(interpreter_information) <> "" AND interpreter_information <> "No Interpreter Used" Then
		If interpreter_language = "English" Then err_msg = err_msg & "~!~" & "1 ^* Language##~##   - Since there is information about interpreter usage, the lanuage should be something other than English. Indicate the language the resident used in the interivew.##~##"
		If trim(interpreter_language) = "" Then err_msg = err_msg & "~!~" & "1 ^* Language##~##   - Since there is information about interpreter usage, enter the language the resident used in the interview in the 'Language' field.##~##"
	End If
	If InStr(UCASE(who_are_we_completing_the_interview_with), "AREP") <> 0 OR InStr(UCASE(who_are_we_completing_the_interview_with), "AUTHORIZED REP") <> 0 Then
		If trim(arep_interview_id_information) = "" Then err_msg = err_msg & "~!~" & "1 ^* Detail AREP Identity Document##~##   - It appears the interview was completed with an AREP (in the field 'Who are you interviewing with?' above). Since identity of the AREP is required if the AREP is the one completing the interview, enter the details about identity of the AREP in the field 'Detail AREP Identity Document'.##~##"
	End If

	If living_situation = "10 - Unknown" OR living_situation = "Blank" or living_situation = "Select" Then err_msg = err_msg & "~!~" & "2 ^* Living Situation?##~##   - Clarify the living situation with the resident for entry."

	For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
            HH_MEMB_ARRAY(imig_status, the_memb) = trim(HH_MEMB_ARRAY(imig_status, the_memb))
    		If HH_MEMB_ARRAY(imig_status, the_memb) <> "" AND HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "" Then err_msg = err_msg & "~!~" & "3 ^* Sponsor?##~##   - Since there is immigration details listed for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ", you need to ask and record if this resident has a sponsor."
    		If HH_MEMB_ARRAY(intend_to_reside_in_mn, the_memb) = "" Then err_msg = err_msg & "~!~" & "3 ^* Intends to Reside in MN##~##   - Indicate if this resident (" & HH_MEMB_ARRAY(full_name_const, the_memb) & ") intends to reside in MN."
    		If the_memb = 0 AND (HH_MEMB_ARRAY(id_verif, the_memb) = "" OR HH_MEMB_ARRAY(id_verif, the_memb) = "NO - No Veer Prvd") Then err_msg = err_msg & "~!~" & "3 ^* Identidty Verification##~##   - Identity is required for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ". Enter the ID information on file/received or indicate that it has been requested."
        End If
	Next

	qual_memb_one = trim(qual_memb_one)
	qual_memb_two = trim(qual_memb_two)
	qual_memb_three = trim(qual_memb_three)
	qual_memb_four = trim(qual_memb_four)
	qual_memb_five = trim(qual_memb_five)
	If qual_question_one = "?" OR (qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type")) Then
		err_msg = err_msg & "~!~" & "5^* Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
		If qual_question_one = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_two = "?" OR (qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type")) Then
		err_msg = err_msg & "~!~" & "5^* Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
		If qual_question_two = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_three = "?" OR (qual_question_three = "Yes" AND (qual_memb_three = "" OR qual_memb_three = "Select or Type")) Then
		err_msg = err_msg & "~!~" & "5^* Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
		If qual_question_three = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_three = "Yes" AND (qual_memb_three = "" OR qual_memb_three = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_four = "?" OR (qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type")) Then
		err_msg = err_msg & "~!~" & "5^* Has anyone in your household been convicted of a drug felony in the past 10 years?"
		If qual_question_four = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_five = "?" OR (qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type")) Then
		err_msg = err_msg & "~!~" & "5^* Is anyone in your household currently violating a condition of parole, probation or supervised release?"
		If qual_question_five = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If

	If expedited_determination_needed = True Then
		If expedited_determination_completed = False Then err_msg = err_msg & "~!~" & "7^* Expedited##~##   - You must complete the process for the Expedited Determination. Press the 'EXPEDITED' button on the right and complete all steps."
	End If

	If err_msg = "" Then interview_questions_clear = TRUE

	If interview_questions_clear = TRUE Then
		'Both signatures - cannot be select or type or blank
		signature_detail = trim(signature_detail)
		second_signature_detail = trim(second_signature_detail)
		signature_person = trim(signature_person)
		second_signature_person = trim(second_signature_person)
		If signature_detail = "Select or Type" OR signature_detail = "" Then err_msg = err_msg & "~!~" & "6^* Signature of Primary Adult##~##   - Indicate how the signature information has been received (or not received)."
		If second_signature_detail = "Select or Type" OR second_signature_detail = "" Then err_msg = err_msg & "~!~" & "6^* Signature of Other Adult##~##   - Indicate how the second signature information has been received (or not received). If no second adult is on the case or the signature of the second adult is not required, select 'Not Required'."
		'If signatires are signed or verbal - then person and date must be completed
		If signature_detail = "Signature Completed" OR signature_detail  = "Accepted Verbally" Then
			If signature_person = "" AND signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & "6^* Signature of Primary Adult - person##~##   - Since the signature was completed, indicate whose sigature it is."
			If IsDate(signature_date) = False Then
				err_msg = err_msg & "~!~" & "6^* Signature of Primary Adult - date##~##   - Enter the date of the signature as a valid date."
			Else
				If DateDiff("d", date, signature_date) > 0 Then err_msg = err_msg & "~!~" & "6^* Signature of Primary Adult - date##~##   - The date of the primary signature cannot be in the future."
			End If
		End If
		If second_signature_detail = "Signature Completed" OR second_signature_detail  = "Accepted Verbally" Then
			If second_signature_person = "" AND second_signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & "6^* Signature of Other Adult - person##~##   - Since the secondary adult signature was completed, indicate whose sigature it is."
			If IsDate(second_signature_date) = False Then
				err_msg = err_msg & "~!~" & "6^* Signature of Other Adult - date##~##   - Enter the date of the signature as a valid date."
			Else
				If DateDiff("d", date, second_signature_date) > 0 Then err_msg = err_msg & "~!~" & "6^* Signature of Other Adult - date##~##   - The date of the primary signature cannot be in the future."
			End If
		End If
		'Interview date must be a date and not in the future
		If IsDate(interview_date) = False Then
			err_msg = err_msg & "~!~" & "6^* Interview Date##~##   - Enter the date of the interview as a valid date."
		Else
			If DateDiff("d", date, interview_date) > 0 Then err_msg = err_msg & "~!~" & "6^* Interview Date##~##   - The date of the interview cannot be in the future."
		End If

		If snap_status = "INACTIVE" AND case_is_expedited = True Then
			If pend_snap_on_case = "?" Then err_msg = err_msg & "~!~11^* SHOULD SNAP BE PENDED ##~##   - Since SNAP is not active on this case, review for possible program eligibility."
		End If
		IF family_cash_case = True OR adult_cash_case = True OR unknown_cash_pending = True Then
			If family_cash_case_yn = "?" Then
				err_msg = err_msg & "~!~11^* IS THIS A FAMILY CASH CASE ##~##   - Since this case has cash active or pending, indicate if this cash is MFIP/DWP."
			ElseIf family_cash_case_yn = "Yes" Then
				If absent_parent_yn = "?" Then err_msg = err_msg & "~!~11^* IS THERE AN ABPS ON THIS CASE ##~##   - Since this is a family cash case, indicate if there is an absent parent for any child on the case."
				If relative_caregiver_yn = "?" Then err_msg = err_msg & "~!~11^* IS THIS A RELATIVE CAREGIVER CASE ##~##   - Since this is a family cash case, indicate if this is a relative caregiver case."
			End If
 		End If
	End If
end function

function define_main_dialog()

	BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"

	  ButtonGroup ButtonPressed
	    If page_display = show_pg_one_memb01_and_exp Then
			Text 497, 17, 60, 10, "INTVW / CAF 1"

			ComboBox 120, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_interview_with, who_are_we_completing_the_interview_with
			ComboBox 120, 30, 75, 45, "Select or Type"+chr(9)+"Phone"+chr(9)+"In Office"+chr(9)+how_are_we_completing_the_interview, how_are_we_completing_the_interview
			EditBox 120, 50, 50, 15, interview_date
			ComboBox 120, 70, 340, 45, "No Interpreter Used"+chr(9)+"Language Line Interpreter Used"+chr(9)+"Interpreter through Henn Co. OMS (Office of Multi-Cultural Services)"+chr(9)+"Interviewer speaks Resident Language"+chr(9)+interpreter_information, interpreter_information
			ComboBox 120, 90, 205, 45, "English"+chr(9)+"Somali"+chr(9)+"Spanish"+chr(9)+"Hmong"+chr(9)+"Russian"+chr(9)+"Oromo"+chr(9)+"Vietnamese"+chr(9)+interpreter_language, interpreter_language
            PushButton 330, 90, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
            EditBox 120, 110, 340, 15, arep_interview_id_information
			EditBox 10, 155, 450, 15, non_applicant_interview_info

		    EditBox 325, 205, 50, 15, exp_q_1_income_this_month
		    EditBox 325, 225, 50, 15, exp_q_2_assets_this_month
		    EditBox 325, 245, 50, 15, exp_q_3_rent_this_month
		    CheckBox 140, 265, 30, 10, "Heat", caf_exp_pay_heat_checkbox
		    CheckBox 175, 265, 65, 10, "Air Conditioning", caf_exp_pay_ac_checkbox
		    CheckBox 245, 265, 45, 10, "Electricity", caf_exp_pay_electricity_checkbox
		    CheckBox 295, 265, 35, 10, "Phone", caf_exp_pay_phone_checkbox
		    CheckBox 340, 265, 35, 10, "None", caf_exp_pay_none_checkbox
		    DropListBox 260, 280, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", exp_migrant_seasonal_formworker_yn
		    DropListBox 380, 295, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", exp_received_previous_assistance_yn
		    EditBox 95, 315, 80, 15, exp_previous_assistance_when
		    EditBox 215, 315, 85, 15, exp_previous_assistance_where
		    EditBox 335, 315, 85, 15, exp_previous_assistance_what
		    DropListBox 175, 335, 40, 45, ""+chr(9)+"No"+chr(9)+"Yes", exp_pregnant_yn
		    ComboBox 270, 335, 150, 45, all_the_clients, exp_pregnant_who

		    Text 10, 15, 110, 10, "Who are you interviewing with?"
			Text 65, 35, 55, 10, "Interview via"
			Text 65, 55, 55, 10, "Interview date"
			Text 30, 75, 85, 10, "Was an Interpreter Used?"
			Text 75, 95, 35, 10, "Language"
			Text 10, 115, 110, 10, "Detail AREP Identity Document"
			Text 120, 130, 300, 10, "(Identity of AREP is required if the interview is being completed with the AREP.)"
			Text 10, 145, 300, 10, "If interview is NOT with a Household Adult, explain relationship and add any details:"

		    GroupBox 25, 185, 400, 170, "CAF 1 Answers - Expedited Section"
			Text 30, 195, 375, 10, "ENTER THE INFORMATION FROM THE CAF HERE."
		    Text 35, 210, 270, 10, "1. How much income (cash or checks) did or will your household get this month?"
		    Text 35, 230, 290, 10, "2. How much does your household (including children) have cash, checking or savings?"
		    Text 35, 250, 225, 10, "3. How much does your household pay for rent/mortgage per month?"
		    Text 45, 265, 90, 10, "What utilities do you pay?"
		    Text 35, 285, 225, 10, "4. Is anyone in your household a migrant or seasonal farm worker?"
		    Text 35, 300, 345, 10, "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
		    Text 45, 320, 50, 10, "If yes, When?"
		    Text 185, 320, 30, 10, "Where?"
		    Text 310, 320, 25, 10, "What?"
		    Text 35, 340, 135, 10, "6. Is anyone in your household pregnant?"
		    Text 225, 340, 45, 10, "If yes, who?"

		ElseIf page_display = show_pg_one_address Then
			Text 504, 32, 60, 10, "CAF ADDR"
			If update_addr = FALSE Then
				Text 70, 55, 305, 15, resi_addr_street_full
				Text 70, 75, 105, 15, resi_addr_city
				Text 205, 75, 110, 45, resi_addr_state
				Text 340, 75, 35, 15, resi_addr_zip
				Text 125, 95, 45, 45, reservation_yn
				Text 245, 85, 130, 15, reservation_name
				Text 125, 115, 45, 45, homeless_yn
				If living_situation = "10 - Unknown" OR living_situation = "Blank" Then
					DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
				Else
					Text 245, 115, 130, 45, living_situation
				End If
				Text 70, 165, 305, 15, mail_addr_street_full
				Text 70, 185, 105, 15, mail_addr_city
				Text 205, 185, 110, 45, mail_addr_state
				Text 340, 185, 35, 15, mail_addr_zip
				Text 20, 240, 90, 15, phone_one_number
				Text 125, 240, 65, 45, phone_one_type
				Text 20, 260, 90, 15, phone_two_number
				Text 125, 260, 65, 45, phone_two_type
				Text 20, 280, 90, 15, phone_three_number
				Text 125, 280, 65, 45, phone_three_type
				Text 325, 220, 50, 15, address_change_date
				Text 255, 255, 120, 45, resi_addr_county
				PushButton 290, 300, 95, 15, "Update Information", update_information_btn
			End If
			If update_addr = TRUE Then
				EditBox 70, 50, 305, 15, resi_addr_street_full
				EditBox 70, 70, 105, 15, resi_addr_city
				DropListBox 205, 70, 110, 45, ""+chr(9)+state_list, resi_addr_state
				EditBox 340, 70, 35, 15, resi_addr_zip
				DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", reservation_yn
				EditBox 245, 90, 130, 15, reservation_name
				DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", homeless_yn
				DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
				EditBox 70, 160, 305, 15, mail_addr_street_full
				EditBox 70, 180, 105, 15, mail_addr_city
				DropListBox 205, 180, 110, 45, ""+chr(9)+state_list, mail_addr_state
				EditBox 340, 180, 35, 15, mail_addr_zip
				EditBox 20, 240, 90, 15, phone_one_number
				DropListBox 125, 240, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_one_type
				EditBox 20, 260, 90, 15, phone_two_number
				DropListBox 125, 260, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_two_type
				EditBox 20, 280, 90, 15, phone_three_number
				DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_three_type
				EditBox 325, 220, 50, 15, address_change_date
				ComboBox 255, 255, 120, 45, county_list+chr(9)+resi_addr_county, resi_addr_county
				PushButton 290, 300, 95, 15, "Save Information", save_information_btn
			End If

			PushButton 325, 145, 50, 10, "CLEAR", clear_mail_addr_btn
			PushButton 205, 240, 35, 10, "CLEAR", clear_phone_one_btn
			PushButton 205, 260, 35, 10, "CLEAR", clear_phone_two_btn
			PushButton 205, 280, 35, 10, "CLEAR", clear_phone_three_btn
			Text 10, 10, 450, 10, "Review the Address informaiton known with the resident. If it needs updating, press the 'Update Information' button to make changes:"
			GroupBox 10, 35, 375, 95, "Residence Address"
			Text 20, 55, 45, 10, "House/Street"
			Text 45, 75, 20, 10, "City"
			Text 185, 75, 20, 10, "State"
			Text 325, 75, 15, 10, "Zip"
			Text 20, 95, 100, 10, "Do you live on a Reservation?"
			Text 180, 95, 60, 10, "If yes, which one?"
			Text 20, 115, 100, 10, "Resident Indicates Homeless:"
			Text 185, 115, 60, 10, "Living Situation?"
			GroupBox 10, 135, 375, 70, "Mailing Address"
			Text 20, 165, 45, 10, "House/Street"
			Text 45, 185, 20, 10, "City"
			Text 185, 185, 20, 10, "State"
			Text 325, 185, 15, 10, "Zip"
			GroupBox 10, 210, 235, 90, "Phone Number"
			Text 20, 225, 50, 10, "Number"
			Text 125, 225, 25, 10, "Type"
			Text 255, 225, 60, 10, "Date of Change:"
			Text 255, 245, 75, 10, "County of Residence:"
		ElseIf page_display = show_pg_memb_list Then
			Text 504, 47, 60, 10, "CAF MEMBs"
			Text 10, 5, 400, 10, "Review information for ALL household members, ensuring the information is accurate."
			Text 10, 15, 400, 10, "You must click on each Person button below and on the left to view each person."

			If update_pers = FALSE Then
				Text 70, 45, 90, 15, HH_MEMB_ARRAY(last_name_const, selected_memb)
				Text 165, 45, 75, 15, HH_MEMB_ARRAY(first_name_const, selected_memb)
				Text 245, 45, 50, 15, HH_MEMB_ARRAY(mid_initial, selected_memb)
				Text 300, 45, 175, 15, HH_MEMB_ARRAY(other_names, selected_memb)
				If HH_MEMB_ARRAY(ssn_verif, selected_memb) = "V - System Verified" Then
					Text 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				Else
					EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				End If
				Text 145, 75, 70, 15, HH_MEMB_ARRAY(date_of_birth, selected_memb)
				Text 220, 75, 50, 45, HH_MEMB_ARRAY(gender, selected_memb)
				Text 275, 75, 90, 45, HH_MEMB_ARRAY(rel_to_applcnt, selected_memb)
				Text 370, 75, 105, 45, HH_MEMB_ARRAY(marital_status, selected_memb)
				Text 70, 105, 110, 15, HH_MEMB_ARRAY(last_grade_completed, selected_memb)
				Text 195, 105, 70, 15, HH_MEMB_ARRAY(mn_entry_date, selected_memb)
				Text 270, 105, 135, 15, HH_MEMB_ARRAY(former_state, selected_memb)
				Text 400, 105, 75, 45, HH_MEMB_ARRAY(citizen, selected_memb)
				Text 70, 135, 60, 45, HH_MEMB_ARRAY(interpreter, selected_memb)
				Text 140, 135, 120, 15, HH_MEMB_ARRAY(spoken_lang, selected_memb)
				Text 140, 165, 120, 15, HH_MEMB_ARRAY(written_lang, selected_memb)
				Text 330, 145, 40, 45, HH_MEMB_ARRAY(ethnicity_yn, selected_memb)
				If the_memb = 0 AND (HH_MEMB_ARRAY(id_verif, the_memb) = "" OR HH_MEMB_ARRAY(id_verif, the_memb) = "NO - No Veer Prvd") Then
					DropListBox 70, 185, 110, 45, ""+chr(9)+id_droplist_info, HH_MEMB_ARRAY(id_verif, selected_memb)
				Else
					Text 70, 185, 110, 10, HH_MEMB_ARRAY(id_verif, selected_memb)
				End If
				PushButton 385, 225, 95, 15, "Update Information", update_information_btn
			End If
			If update_pers = TRUE Then
				EditBox 70, 45, 90, 15, HH_MEMB_ARRAY(last_name_const, selected_memb)
				EditBox 165, 45, 75, 15, HH_MEMB_ARRAY(first_name_const, selected_memb)
				EditBox 245, 45, 50, 15, HH_MEMB_ARRAY(mid_initial, selected_memb)
				EditBox 300, 45, 175, 15, HH_MEMB_ARRAY(other_names, selected_memb)
				EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				EditBox 145, 75, 70, 15, HH_MEMB_ARRAY(date_of_birth, selected_memb)
				DropListBox 220, 75, 50, 45, ""+chr(9)+"Male"+chr(9)+"Female", HH_MEMB_ARRAY(gender, selected_memb)
				DropListBox 275, 75, 90, 45, memb_panel_relationship_list, HH_MEMB_ARRAY(rel_to_applcnt, selected_memb)
				DropListBox 370, 75, 105, 45, marital_status_list, HH_MEMB_ARRAY(marital_status, selected_memb)
				EditBox 70, 105, 110, 15, HH_MEMB_ARRAY(last_grade_completed, selected_memb)
				EditBox 185, 105, 70, 15, HH_MEMB_ARRAY(mn_entry_date, selected_memb)
				EditBox 260, 105, 135, 15, HH_MEMB_ARRAY(former_state, selected_memb)
				DropListBox 400, 105, 75, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(citizen, selected_memb)
				DropListBox 70, 135, 60, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(interpreter, selected_memb)
				EditBox 140, 135, 120, 15, HH_MEMB_ARRAY(spoken_lang, selected_memb)
				EditBox 140, 165, 120, 15, HH_MEMB_ARRAY(written_lang, selected_memb)
				DropListBox 330, 145, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(ethnicity_yn, selected_memb)
				DropListBox 70, 185, 110, 45, ""+chr(9)+id_droplist_info, HH_MEMB_ARRAY(id_verif, selected_memb)

				PushButton 385, 225, 95, 15, "Save Information", save_information_btn
			End If
			CheckBox 330, 170, 30, 10, "Asian", HH_MEMB_ARRAY(race_a_checkbox, selected_memb)
			CheckBox 330, 180, 30, 10, "Black", HH_MEMB_ARRAY(race_b_checkbox, selected_memb)
			CheckBox 330, 190, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(race_n_checkbox, selected_memb)
			CheckBox 330, 200, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(race_p_checkbox, selected_memb)
			CheckBox 330, 210, 130, 10, "White", HH_MEMB_ARRAY(race_w_checkbox, selected_memb)
			CheckBox 70, 210, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(snap_req_checkbox, selected_memb)
			CheckBox 125, 210, 65, 10, "Cash programs", HH_MEMB_ARRAY(cash_req_checkbox, selected_memb)
			CheckBox 195, 210, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(emer_req_checkbox, selected_memb)
			CheckBox 280, 210, 30, 10, "NONE", HH_MEMB_ARRAY(none_req_checkbox, selected_memb)
			If selected_memb = 0 Then
				DropListBox 70, 265, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb)
			Else
				DropListBox 70, 265, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Not in HH", HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb)
			End If
			EditBox 155, 265, 205, 15, HH_MEMB_ARRAY(imig_status, selected_memb)
			DropListBox 365, 265, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(clt_has_sponsor, selected_memb)
			EditBox 70, 290, 405, 15, HH_MEMB_ARRAY(client_notes, selected_memb)
			If HH_MEMB_ARRAY(ref_number, selected_memb) = "" Then
				GroupBox 65, 25, 415, 200, "Person " & selected_memb+1
				GroupBox 65, 245, 415, 100, "Person " & selected_memb+1 & "  ---  Interview Questions"
			Else
				GroupBox 65, 25, 415, 200, "Person " & selected_memb+1 & " - MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb)
				GroupBox 65, 245, 415, 100, "Person " & selected_memb+1 & " - MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb) & "  ---  Interview Questions"

			End If
			y_pos = 35
			For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
				If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
                    If the_memb = selected_memb Then
    					Text 20, y_pos + 1, 45, 10, "Person " & (the_memb + 1)
    				Else
    					PushButton 10, y_pos, 45, 10, "Person " & (the_memb + 1), HH_MEMB_ARRAY(button_one, the_memb)
    				End If
    				y_pos = y_pos + 10
                End If
			Next
            If HH_MEMB_ARRAY(pers_in_maxis, selected_memb) = False Then PushButton 375, 30, 105, 13, "Remove Member from Script", HH_MEMB_ARRAY(button_two, selected_memb)
			y_pos = y_pos + 10
			PushButton 10, 335, 45, 10, "Add Person", add_person_btn
			Text 70, 35, 50, 10, "Last Name"
			Text 165, 35, 50, 10, "First Name"
			Text 245, 35, 50, 10, "Middle Name"
			Text 300, 35, 50, 10, "Other Names"
			Text 70, 65, 55, 10, "Soc Sec Number"
			Text 145, 65, 45, 10, "Date of Birth"
			Text 220, 65, 45, 10, "Gender"
			Text 275, 65, 90, 10, "Relationship to MEMB 01"
			Text 370, 65, 50, 10, "Marital Status"
			Text 70, 95, 75, 10, "Last Grade Completed"
			Text 185, 95, 55, 10, "Moved to MN on"
			Text 260, 95, 65, 10, "Moved to MN from"
			Text 400, 95, 75, 10, "US Citizen or National"
			Text 70, 125, 40, 10, "Interpreter?"
			Text 140, 125, 95, 10, "Preferred Spoken Language"
			Text 140, 155, 95, 10, "Preferred Written Language"
			Text 70, 175, 65, 10, "Identity Verification"
			GroupBox 325, 125, 155, 100, "Demographics"
			Text 330, 135, 35, 10, "Hispanic?"
			Text 330, 160, 50, 10, "Race"
			Text 70, 200, 145, 10, "Which programs is this person requesting?"
			Text 70, 255, 80, 10, "Intends to reside in MN"
			Text 155, 255, 65, 10, "Immigration Status"
			Text 365, 255, 50, 10, "Sponsor?"
			Text 70, 280, 50, 10, "Notes:"
		ElseIf page_display = show_case_info Then																	' THIS SECTION IS THE VARIATION FROM THE FULL INTERVIEW SCRIPT
			Text 505, 62, 60, 10, "DETAILS"
			GroupBox 5, 5, 475, 255, "Case Information"
			Text 10, 20, 190, 10, "Household Information (School, Disability, Absense):"
			EditBox 10, 30, 460, 15, household_info
			Text 10, 55, 195, 10, "Earned Income (Jobs, Self-Employment, start or stop work):"
			EditBox 10, 65, 460, 15, earned_income_info
			Text 10, 90, 255, 10, "Unearned Income (Social Security, Child Support, Unemployment, VA, etc):"
			EditBox 10, 100, 460, 15, unearned_income_info
			Text 10, 125, 245, 10, "Housing Expenses (Rent, Mortgage, Subsidy, Utilities):"
			EditBox 10, 135, 460, 15, housing_expanese_info
			Text 10, 160, 330, 10, "Other Expenses (Child Support Payments, Child/Adult Care Expenses, Medical Expesnse, etc.):"
			EditBox 10, 170, 460, 15, other_expenses_info
			Text 10, 195, 245, 10, "Assets (Accounts, Real Estate, Vehicles, Securities, etc):"
			EditBox 10, 205, 460, 15, assets_info
			Text 10, 230, 245, 10, "Other Notes (Unique Case Scenarios, Changes, Additional Details, etc.):"
			EditBox 10, 240, 460, 15, other_notes_info
			GroupBox 5, 265, 475, 95, "Interview Specifics"
			Text 10, 280, 245, 10, "Detail any Verbal Changes from the FORM:"
			EditBox 10, 290, 460, 15, form_changes_info
			Text 10, 310, 250, 10, "Check ONLY the boxes of information you discussed (not all are required):"
			CheckBox 15, 320, 110, 10, "Rights and Responsibilities", rights_responsibilities_checkbox
			CheckBox 15, 330, 110, 10, "Privacy Practices", privacy_practices_checkbox
			CheckBox 15, 340, 120, 10, "EBT Card Process and Information", ebt_card_checkbox
			CheckBox 140, 320, 110, 10, "Complaints and Civil Rights", complaints_civil_rights_checkbox
			CheckBox 140, 330, 110, 10, "Reporting Responsibilities", reporting_resp_checkbox
			CheckBox 140, 340, 110, 10, "Program Information", program_info_checkbox
			CheckBox 250, 320, 120, 10, "Child Support Referral and Coop", child_support_info_checkbox
			CheckBox 250, 330, 120, 10, "MFIP Minor Caregiver Information", mfip_minor_caregiver_checkbox
			CheckBox 250, 340, 120, 10, "Renewal Process and Information", renewal_info_checkbox
			CheckBox 380, 320, 110, 10, "IEVS Information", ievs_info_checkbox
			CheckBox 380, 330, 110, 10, "Appeal Rights", appeal_rights_checkbox
		ElseIf page_display = show_qual Then
			Text 500, 77, 60, 10, "CAF QUAL Q"

			DropListBox 220, 40, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_one
			ComboBox 340, 40, 105, 45, all_the_clients, qual_memb_one
			DropListBox 220, 80, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_two
			ComboBox 340, 80, 105, 45, all_the_clients, qual_memb_two
			DropListBox 220, 110, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_three
			ComboBox 340, 110, 105, 45, all_the_clients, qual_memb_three
			DropListBox 220, 140, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_four
			ComboBox 340, 140, 105, 45, all_the_clients, qual_memb_four
			DropListBox 220, 160, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_five
			ComboBox 340, 160, 105, 45, all_the_clients, qual_memb_five

			Text 10, 10, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the resident. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
			Text 10, 40, 200, 40, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			Text 10, 80, 195, 30, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			Text 10, 110, 195, 30, "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			Text 10, 140, 195, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
			Text 10, 160, 195, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			Text 260, 40, 70, 10, "Household Member:"
			Text 260, 80, 70, 10, "Household Member:"
			Text 260, 110, 70, 10, "Household Member:"
			Text 260, 140, 70, 10, "Household Member:"
			Text 260, 160, 70, 10, "Household Member:"
		ElseIf page_display = show_pg_last Then
			Text 498, 92, 60, 10, "CAF Last Page"

			GroupBox 5, 5, 475, 60, "Confirm Authorized Representative"

			If arep_exists =  False Then Text 15, 25, 300, 10, "There is no Authorized Representative"
			If arep_exists = True Then
				Text 10, 20, 175, 10, "AREP Name: " & arep_name
				Text 150, 20, 125, 10, "Relationship: " & arep_relationship
				Text 275, 20, 100, 10, "Phone Number: " & arep_phone_number
				Text 10, 35, 385, 10, "Address: " & arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip
				' Text 85, 45, 385, 10, arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip
				CheckBox 20, 50, 55, 10, "Fill out forms", arep_complete_forms_checkbox
				CheckBox 80, 50, 50, 10, "Get notices", arep_get_notices_checkbox
				CheckBox 135, 50, 140, 10, "Get and use my SNAP benefit", arep_use_SNAP_checkbox
				' Text 20, 60, 50, 10, "SNAP benefits"
			End If
			PushButton 390, 47, 85, 13, "Update AREP Detail", update_information_btn

			' (less 35)
		    GroupBox 5, 70, 475, 75, "Signatures"
		    Text 10, 85, 90, 10, "Signature of Primary Adult"
		    ComboBox 105, 80, 110, 45, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+"Not Required"+chr(9)+signature_detail, signature_detail
		    Text 220, 85, 25, 10, "person"
		    ComboBox 250, 80, 115, 45, all_the_clients+chr(9)+signature_person, signature_person
		    Text 375, 85, 20, 10, "date"
		    EditBox 400, 80, 50, 15, signature_date
		    Text 10, 105, 90, 10, "Signature of Other Adult"
		    ComboBox 105, 100, 110, 45, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Not Required"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+second_signature_detail, second_signature_detail
		    Text 220, 105, 25, 10, "person"
		    ComboBox 250, 100, 115, 45, all_the_clients+chr(9)+second_signature_person, second_signature_person
		    Text 375, 105, 20, 10, "date"
		    EditBox 400, 100, 50, 15, second_signature_date

			Text 10, 125, 130, 10, "Resident signature accepted verbally?"
			DropListBox 135, 120, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_verbally_yn
			Text 335, 125, 50, 10, "Interview Date:"
			EditBox 390, 120, 60, 15, interview_date

			GroupBox 5, 150, 475, 200, "Benefit Detail"
			y_pos = 165
			If interview_questions_clear = False Then
				Text 15, 165, 450, 10, "ADDITIONAL QUESTIONS BEFORE ASSESMENT IS COMPLETE."
				y_pos = 185
			End If
			' appears_expedited
			' expedited_delay_info
			If cash_request = True Then
				If the_process_for_cash = "Renewal" Then Text 15, y_pos, 450, 10, "CASH Case at " & the_process_for_cash & " for " & next_cash_revw_mo & "/" & next_cash_revw_yr
				If the_process_for_cash = "Application" Then Text 15, y_pos, 450, 10, "CASH Case at " & the_process_for_cash
				y_pos = y_pos + 15
			End If
			If snap_request = True Then
				Text 15, y_pos, 450, 10, "SNAP is active on this case - Expedited Determination not needed."
				If the_process_for_snap = "Renewal" Then Text 15, y_pos, 450, 10, "SNAP Case at " & the_process_for_snap & " for " & next_snap_revw_mo & "/" & next_snap_revw_yr
				If the_process_for_snap = "Application" Then Text 15, y_pos, 450, 10, "SNAP Case at " & the_process_for_snap
				y_pos = y_pos + 15
			End If
			If emer_request = True Then
				Text 15, y_pos, 450, 10, "EMERGENCY Request on Case is " & type_of_emer
				y_pos = y_pos + 15
			End If
			If expedited_determination_needed = True Then
				If expedited_determination_completed = False Then
					Text 15, y_pos, 450, 10, "COMPLETE THE EXPEDITED DETERMINATION - press the button 'EXPEDITED' on the right."
					y_pos = y_pos + 15
				Else
					If case_assesment_text <> "" Then
						Text 15, y_pos, 450, 10, case_assesment_text
						y_pos = y_pos + 10
					End If

					If next_steps_one <> "" Then
						Text 20, y_pos, 450, 20, next_steps_one
						y_pos = y_pos + 20
					End If
					If next_steps_two <> "" Then
						Text 20, y_pos, 450, 20, next_steps_two
						y_pos = y_pos + 20
					End If
					If next_steps_three <> "" Then
						Text 20, y_pos, 450, 20, next_steps_three
						y_pos = y_pos + 20
					End If
					If next_steps_four <> "" Then
						Text 20, y_pos, 450, 20, next_steps_four
						y_pos = y_pos + 20
					End If
				End If
			End If

			IF family_cash_case = True OR adult_cash_case = True OR unknown_cash_pending = True Then
				Text 15, y_pos, 100, 10, "Is this a Family Cash case?"
				DropListBox 115, y_pos - 5, 50, 45, "?"+chr(9)+"Yes"+chr(9)+"No", family_cash_case_yn
				y_pos = y_pos + 20
				If family_cash_case_yn = "?" OR family_cash_case_yn = "Yes" Then
					Text 15, y_pos, 175, 10, "Is there an Absent Parent for any children on this case?"
					DropListBox 190, y_pos - 5, 50, 45, "?"+chr(9)+"Yes"+chr(9)+"No", absent_parent_yn
					Text 255, y_pos, 115, 10, "Is this a relative caregiver case?"
					DropListBox 370, y_pos - 5, 50, 45, "?"+chr(9)+"Yes"+chr(9)+"No", relative_caregiver_yn
					y_pos = y_pos + 20

					Text 15, y_pos, 150, 10, "Are there any minor caregivers on this case?"
					DropListBox 165, y_pos - 5, 135, 45, "No - all cargivers are over 20"+chr(9)+"Yes - Caregiver is 18 - 20 years old"+chr(9)+"Yes - Caregiver is under 18", minor_caregiver_yn
					y_pos = y_pos + 20
				End If

			End If

			' THIS SECTION IS NORMALLY AT THE END OF THE INTERVIEW SCRIPT BUT MOVED HERE FOR SUMMARY
		  	grp_len = 20
			grp_box_y_pos = y_pos
			y_pos = y_pos + 15
			If verifs_selected <> "" Then

				verifs_selected = trim(verifs_selected)
				If right(verifs_selected, 1) = ";" Then
					verifs_to_view = left(verifs_selected, len(verifs_selected)-1)
				Else
					verifs_to_view = verifs_selected
				End If

				If verifs_to_view <> "" Then
					array_of_verifs_selected = ""
					If InStr(verifs_to_view, ";") = 0 Then
						array_of_verifs_needed = array(verifs_to_view)
					Else
						array_of_verifs_needed = split(verifs_to_view, ";")
					End If

					for each verif_item in array_of_verifs_needed
						Text 17, y_pos, 445, 10, trim(verif_item)
						y_pos = y_pos + 10
						grp_len = grp_len + 10
					next
				End If
			Else
				Text 10, y_pos, 500, 10, "NO VERIFICATIONS HAVE BEEN LISTED YET"
				grp_len = grp_len + 15
			End If
			GroupBox 10, grp_box_y_pos, 455, grp_len, "Verifications Entered"

	    ElseIf page_display = show_arep_page Then
			If arep_addr_state = "" Then arep_addr_state = "MN Minnesota"
			If CAF_arep_addr_state = "" Then CAF_arep_addr_state = "MN Minnesota"
			' GroupBox 5, 5, 475, 300, "Authorized Representative Detail"

			If arep_in_MAXIS = True AND MAXIS_arep_updated = False Then
				GroupBox 5, 5, 475, 140, "AREP from MAXIS"
				Text 10, 20, 45, 10, "AREP Name"
				EditBox 10, 30, 170, 15, arep_name
				Text 185, 20, 50, 10, "Relationship"
				ComboBox 185, 30, 120, 45, "Select or Type"+chr(9)+"Parent"+chr(9)+"Grandparent"+chr(9)+"Child"+chr(9)+"Grandchild"+chr(9)+"Aunt/Uncle"+chr(9)+"Neice/Nephew"+chr(9)+"Caretaker"+chr(9)+"Unrelated"+chr(9)+arep_relationship, arep_relationship
				Text 310, 20, 50, 10, "Phone Number"
				EditBox 310, 30, 85, 15, arep_phone_number
				Text 10, 50, 35, 10, "Address"
				EditBox 10, 60, 170, 15, arep_addr_street
				Text 185, 50, 25, 10, "City"
				EditBox 185, 60, 85, 15, arep_addr_city
				Text 275, 50, 25, 10, "State"
				DropListBox 275, 60, 65, 45, state_list, arep_addr_state
				Text 345, 50, 35, 10, "Zip Code"
				EditBox 345, 60, 50, 15, arep_addr_zip

				CheckBox 20, 80, 55, 10, "Fill out forms", arep_complete_forms_checkbox
				CheckBox 80, 80, 50, 10, "Get notices", arep_get_notices_checkbox
				CheckBox 135, 80, 140, 10, "Get and use my SNAP benefit", arep_use_SNAP_checkbox

				GroupBox 20, 95, 460, 50, "Actions to Take on this AREP Information"
				CheckBox 30, 110, 250, 10, "Check Here if this AREP is ALSO Listed as an AREP on the CAF", arep_on_CAF_checkbox
				Text 30, 130, 165, 10, "Does the Resident want this AREP to Continue?"
				DropListBox 195, 125, 150, 15, "Select One..."+chr(9)+"Yes - keep this AREP"+chr(9)+"No - remove this AREP from my case", arep_action
			ElseIf arep_in_MAXIS = True AND MAXIS_arep_updated = True Then
				GroupBox 5, 5, 475, 140, "AREP Updated or Entered into Script"
				Text 10, 20, 45, 10, "AREP Name"
				EditBox 10, 30, 170, 15, arep_name
				Text 185, 20, 50, 10, "Relationship"
				ComboBox 185, 30, 120, 45, "Select or Type"+chr(9)+"Parent"+chr(9)+"Grandparent"+chr(9)+"Child"+chr(9)+"Grandchild"+chr(9)+"Aunt/Uncle"+chr(9)+"Neice/Nephew"+chr(9)+"Caretaker"+chr(9)+"Unrelated"+chr(9)+arep_relationship, arep_relationship
				Text 310, 20, 50, 10, "Phone Number"
				EditBox 310, 30, 85, 15, arep_phone_number
				Text 10, 50, 35, 10, "Address"
				EditBox 10, 60, 170, 15, arep_addr_street
				Text 185, 50, 25, 10, "City"
				EditBox 185, 60, 85, 15, arep_addr_city
				Text 275, 50, 25, 10, "State"
				DropListBox 275, 60, 65, 45, state_list, arep_addr_state
				Text 345, 50, 35, 10, "Zip Code"
				EditBox 345, 60, 50, 15, arep_addr_zip

				CheckBox 20, 80, 55, 10, "Fill out forms", arep_complete_forms_checkbox
				CheckBox 80, 80, 50, 10, "Get notices", arep_get_notices_checkbox
				CheckBox 135, 80, 140, 10, "Get and use my SNAP benefit", arep_use_SNAP_checkbox

				GroupBox 20, 95, 460, 50, "Actions to Take on this AREP Information"
				CheckBox 30, 110, 250, 10, "Check Here if this AREP is ALSO Listed as an AREP on the CAF", arep_on_CAF_checkbox
				Text 30, 130, 165, 10, "Does the Resident want this AREP to Continue?"
				DropListBox 195, 125, 150, 15, "Select One..."+chr(9)+"Yes - keep this AREP"+chr(9)+"No - remove this AREP from my case", arep_action
			ElseIf arep_in_MAXIS = False Then
				GroupBox 5, 5, 475, 140, "AREP reported Verbally"
				Text 10, 20, 45, 10, "AREP Name"
				EditBox 10, 30, 170, 15, arep_name
				Text 185, 20, 50, 10, "Relationship"
				ComboBox 185, 30, 120, 45, "Select or Type"+chr(9)+"Parent"+chr(9)+"Grandparent"+chr(9)+"Child"+chr(9)+"Grandchild"+chr(9)+"Aunt/Uncle"+chr(9)+"Neice/Nephew"+chr(9)+"Caretaker"+chr(9)+"Unrelated"+chr(9)+arep_relationship, arep_relationship
				Text 310, 20, 50, 10, "Phone Number"
				EditBox 310, 30, 85, 15, arep_phone_number
				Text 10, 50, 35, 10, "Address"
				EditBox 10, 60, 170, 15, arep_addr_street
				Text 185, 50, 25, 10, "City"
				EditBox 185, 60, 85, 15, arep_addr_city
				Text 275, 50, 25, 10, "State"
				DropListBox 275, 60, 65, 45, state_list, arep_addr_state
				Text 345, 50, 35, 10, "Zip Code"
				EditBox 345, 60, 50, 15, arep_addr_zip

				CheckBox 20, 80, 55, 10, "Fill out forms", arep_complete_forms_checkbox
				CheckBox 80, 80, 50, 10, "Get notices", arep_get_notices_checkbox
				CheckBox 135, 80, 140, 10, "Get and use my SNAP benefit", arep_use_SNAP_checkbox

				GroupBox 20, 95, 460, 50, "Actions to Take on this AREP Information"
				CheckBox 30, 110, 250, 10, "Check Here if this AREP is ALSO Listed as an AREP on the CAF", arep_on_CAF_checkbox
				Text 30, 130, 165, 10, "Does the Resident want this AREP to Continue?"
				DropListBox 195, 125, 150, 15, "Select One..."+chr(9)+"Yes - keep this AREP"+chr(9)+"No - remove this AREP from my case", arep_action

			End If

			GroupBox 5, 160, 475, 125, "AREP on CAF"
			Text 10, 175, 45, 10, "AREP Name"
			EditBox 10, 185, 170, 15, CAF_arep_name
			Text 185, 175, 50, 10, "Relationship"
			ComboBox 185, 185, 120, 45, "Select or Type"+chr(9)+"Parent"+chr(9)+"Grandparent"+chr(9)+"Child"+chr(9)+"Grandchild"+chr(9)+"Aunt/Uncle"+chr(9)+"Neice/Nephew"+chr(9)+"Caretaker"+chr(9)+"Unrelated"+chr(9)+CAF_arep_relationship, CAF_arep_relationship
			Text 310, 175, 50, 10, "Phone Number"
			EditBox 310, 185, 85, 15, CAF_arep_phone_number
			Text 10, 205, 35, 10, "Address"
			EditBox 10, 215, 170, 15, CAF_arep_addr_street
			Text 185, 205, 25, 10, "City"
			EditBox 185, 215, 85, 15, CAF_arep_addr_city
			Text 275, 205, 25, 10, "State"
			DropListBox 275, 215, 65, 45, state_list, CAF_arep_addr_state
			Text 345, 205, 35, 10, "Zip Code"
			EditBox 345, 215, 50, 15, CAF_arep_addr_zip

			CheckBox 20, 235, 55, 10, "Fill out forms", CAF_arep_complete_forms_checkbox
			CheckBox 80, 235, 50, 10, "Get notices", CAF_arep_get_notices_checkbox
			CheckBox 135, 235, 140, 10, "Get and use my SNAP benefit", CAF_arep_use_SNAP_checkbox

			GroupBox 20, 250, 460, 35, "Actions to Take on this AREP Information"
			Text 30, 270, 175, 10, "Does the Resident want this AREP added to the Case?"
			DropListBox 210, 265, 150, 15, "Select One..."+chr(9)+"Yes - add to MAXIS"+chr(9)+"No - do not allow this AREP", CAF_arep_action
			' CheckBox 30, 285, 200, 10, "Check Here if this AREP is ALSO Listed on the CAF", CAF_arep_on_CAF_checkbox

			Text 10, 295, 85, 10, "Authorization of AREP:"
			DropListBox 95, 290, 175, 15, "Select One..."+chr(9)+"AREP authorized verbal"+chr(9)+"AREP Authorized by entry on the CAF"+chr(9)+"AREP authorized by seperate written document"+chr(9)+"AREP previously entered - authorization unknown"+chr(9)+"DO NOT AUTHORIZE AN AREP"+chr(9)+arep_authorization, arep_authorization
			PushButton 395, 292, 85, 13, "Save AREP Detail", save_information_btn

		ElseIf page_display = expedited_determination Then
			btn_pos = 180
			If discrepancies_exist = True Then btn_pos = btn_pos + 15
			Text 505, btn_pos+2, 60, 10, "EXPEDITED"

		End If

		Text 485, 5, 75, 10, "---   DIALOGS   ---"
		Text 485, 17, 10, 10, "1"
		Text 485, 32, 10, 10, "2"
		Text 485, 47, 10, 10, "3"
		Text 485, 62, 10, 10, "4"
		Text 485, 77, 10, 10, "5"
		Text 485, 92, 10, 10, "6"
		If page_display <> show_pg_one_memb01_and_exp 	Then PushButton 495, 15, 55, 13, "INTVW / CAF 1", caf_page_one_btn
		If page_display <> show_pg_one_address 			Then PushButton 495, 30, 55, 13, "CAF ADDR", caf_addr_btn
		If page_display <> show_pg_memb_list 			Then PushButton 495, 45, 55, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_case_info				Then PushButton 495, 60, 55, 13, "DETAILS", case_details_btn
		If page_display <> show_qual 					Then PushButton 495, 75, 55, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last 				Then PushButton 495, 90, 55, 13, "CAF Last Page", caf_last_page_btn
		btn_pos = 105
		If expedited_determination_needed = True Then
			Text 485, btn_pos + 2, 10, 10, "7"
			If page_display <> expedited_determination Then PushButton 495, btn_pos, 55, 13, "EXPEDITED", expedited_determination_btn
			btn_pos = btn_pos + 15
		End If
		PushButton 10, 365, 130, 15, "View Verifications", verif_button
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn

	EndDialog

end function

function dialog_movement()
	If page_display = show_arep_page Then
		arep_exists = True
		If arep_in_MAXIS = True Then
			If arep_name <> MAXIS_arep_name Then MAXIS_arep_updated = True
			' If arep_relationship <> MAXIS_arep_relationship Then MAXIS_arep_updated = True
			If arep_phone_number <> MAXIS_arep_phone_number Then MAXIS_arep_updated = True
			If arep_addr_street <> MAXIS_arep_addr_street Then MAXIS_arep_updated = True
			If arep_addr_city <> MAXIS_arep_addr_city Then MAXIS_arep_updated = True
			If arep_addr_state <> MAXIS_arep_addr_state Then MAXIS_arep_updated = True
			If arep_addr_zip <> MAXIS_arep_addr_zip Then MAXIS_arep_updated = True

		End If
		If arep_on_CAF_checkbox = checked Then
			CAF_arep_name = arep_name
			CAF_arep_relationship = arep_relationship
			CAF_arep_phone_number = arep_phone_number
			CAF_arep_addr_street = arep_addr_street
			CAF_arep_addr_city = arep_addr_city
			CAF_arep_addr_state = arep_addr_state
			CAF_arep_addr_zip = arep_addr_zip

			CAF_arep_complete_forms_checkbox = arep_complete_forms_checkbox
			CAF_arep_get_notices_checkbox = arep_get_notices_checkbox
			CAF_arep_use_SNAP_checkbox = arep_use_SNAP_checkbox

			If arep_action = "Yes - keep this AREP" Then CAF_arep_action = "Yes - add to MAXIS"
			If arep_action = "No - remove this AREP from my case" Then CAF_arep_action = "No - do not allow this AREP"
		End If

		If arep_on_CAF_checkbox = checked OR trim(CAF_arep_name) <> "" Then arep_authorization = "AREP Authorized by entry on the CAF"
		If arep_authorization = "DO NOT AUTHORIZE AN AREP" Then
			arep_action = "No - remove this AREP from my case"
			CAF_arep_action = "No - do not allow this AREP"
			arep_exists = False
			arep_authorized = False
		End If
		If CAF_arep_name = "" AND arep_name = "" Then
			arep_authorization = ""
			arep_action = ""
			CAF_arep_action = ""
			arep_exists = False
		End If
		If arep_authorization <> "" AND arep_authorization <> "Select One..." and arep_exists = True Then arep_authorized = True

	End If
	arep_and_CAF_arep_match = False
	If CAF_arep_name = arep_name Then arep_and_CAF_arep_match = True


	For i = 0 to Ubound(HH_MEMB_ARRAY, 2)
		' If HH_MEMB_ARRAY(i).imig_exists = TRUE Then case_has_imig = TRUE
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = HH_MEMB_ARRAY(button_one, i) Then
			If page_display = show_pg_memb_list Then selected_memb = i
		End If
        If ButtonPressed = HH_MEMB_ARRAY(button_two, i) Then
            HH_MEMB_ARRAY(ignore_person, i) = True
            selected_memb = 0
        End If
	Next

	If ButtonPressed = open_hsr_manual_transfer_page_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Case-Transfers.aspx"

	If arep_name <> "" Then arep_exists = True
	If ButtonPressed = update_information_btn Then
		If page_display = show_pg_one_address Then update_addr = TRUE
		If page_display = show_pg_memb_list Then update_pers = TRUE
		If page_display = show_pg_last Then page_display = show_arep_page
		' MsgBox update_arep & " - in dlg move"
	End If
	If ButtonPressed = save_information_btn Then
		If page_display = show_pg_one_address Then update_addr = FALSE
		If page_display = show_pg_memb_list Then update_pers = FALSE
		If page_display = show_arep_page Then page_display = show_pg_last

	End If
	If ButtonPressed = clear_mail_addr_btn Then
		' phone_one_number = ""
		' phone_one_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_one_btn Then
		phone_one_number = ""
		phone_one_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_two_btn Then
		phone_two_number = ""
		phone_two_type = "Select One..."
	End If
	If ButtonPressed = clear_phone_three_btn Then
		phone_three_number = ""
		phone_three_type = "Select One..."
	End If

	If page_display = show_pg_memb_info AND ButtonPressed = -1 Then ButtonPressed = next_memb_btn

	If ButtonPressed = next_memb_btn Then
		Do
            memb_selected = memb_selected + 1
            If HH_MEMB_ARRAY(ignore_person, memb_selected) = True Then memb_selected = memb_selected + 1
        Loop until HH_MEMB_ARRAY(ignore_person, memb_selected) = False OR memb_selected > UBound(HH_MEMB_ARRAY, 2)
		If memb_selected > UBound(HH_MEMB_ARRAY, 2) Then ButtonPressed = next_btn
	End If

	If ButtonPressed = add_person_btn Then
		last_clt = UBound(HH_MEMB_ARRAY, 2)
		new_clt = last_clt + 1
		ReDim Preserve HH_MEMB_ARRAY(last_const, new_clt)
		HH_MEMB_ARRAY(button_one, new_clt) = 500 + new_clt
		HH_MEMB_ARRAY(button_two, new_clt) = 600 + new_clt
        HH_MEMB_ARRAY(pers_in_maxis, new_clt) = False

		selected_memb = new_clt
		update_pers = TRUE
	End If
	If ButtonPressed = exp_income_guidance_btn Then
		call guide_through_app_month_income
	End If
	If ButtonPressed = -1 Then ButtonPressed = next_btn
	If ButtonPressed = next_btn Then
		If page_display = show_pg_one_memb01_and_exp 	Then ButtonPressed = caf_addr_btn
		If page_display = show_pg_one_address 			Then ButtonPressed = caf_membs_btn
		If page_display = show_pg_memb_list 			Then ButtonPressed = case_details_btn
		If page_display = show_case_info 				Then ButtonPressed = caf_qual_q_btn
		If page_display = show_qual 					Then ButtonPressed = caf_last_page_btn
		If page_display = show_pg_last 					Then ButtonPressed = finish_interview_btn
		If expedited_determination_needed = True Then
			If expedited_determination_completed = False Then
				ButtonPressed = expedited_determination_btn
			Else
				ButtonPressed = finish_interview_btn
			End If
		End If
	End If

	If ButtonPressed = caf_page_one_btn Then
		page_display = show_pg_one_memb01_and_exp
	End If
	If ButtonPressed = caf_addr_btn Then
		page_display = show_pg_one_address
	End If
	If ButtonPressed = caf_membs_btn Then
		page_display = show_pg_memb_list
	End If
	If ButtonPressed = case_details_btn Then
		page_display = show_case_info
	End If
	If ButtonPressed = caf_qual_q_btn Then
		page_display = show_qual
	End If
	If ButtonPressed = caf_last_page_btn Then
		page_display = show_pg_last
	End If
	If ButtonPressed = expedited_determination_btn Then
		' page_display = expedited_determination
		STATS_manualtime = STATS_manualtime + 150
		call display_expedited_dialog
	End If

	If ButtonPressed = finish_interview_btn Then leave_loop = TRUE
	If ButtonPressed > 10000 Then
		save_button = ButtonPressed
		If ButtonPressed = page_1_step_1_btn Then call explain_dialog_actions("PAGE 1", "STEP 1")
		If ButtonPressed = page_1_step_2_btn Then call explain_dialog_actions("PAGE 1", "STEP 2")

		ButtonPressed = save_button
	End If
end function

function display_errors(the_err_msg, execute_nav, show_err_msg_during_movement)
    If the_err_msg <> "" Then       'If the error message is blank - there is nothing to show.
        If left(the_err_msg, 3) = "~!~" Then the_err_msg = right(the_err_msg, len(the_err_msg) - 3)     'Trimming the message so we don't have a blank array item
        err_array = split(the_err_msg, "~!~")           'making the list of errors an array.

        error_message = ""                              'blanking out variables
        msg_header = ""
        for each message in err_array                   'going through each error message to order them and add headers'
			If show_err_msg_during_movement = False OR ButtonPressed = finish_interview_btn Then
	            current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing = trim(current_listing)
	            If current_listing <> msg_header Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
	                If current_listing = "1"  Then tagline = ": INTVW / CAF 1"        'Adding a specific tagline to the header for the errors
	                If current_listing = "2"  Then tagline = ": CAF ADDR"
	                If current_listing = "3"  Then tagline = ": CAF MEMBs"
					If current_listing = "4"  Then tagline = ": DETAILS"
					If current_listing = "5" Then tagline = ": CAF QUAL Q"
	                If current_listing = "6" Then tagline = ": CAF Last Page"
					If current_listing = "7" Then tagline = ": Clarifications"
	                error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
	            End If
	            if msg_header = "" Then back_to_dialog = current_listing
	            msg_header = current_listing        'setting for the next loop

	            message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

	            error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
			ElseIf show_err_msg_during_movement = TRUE Then
				If page_display = show_pg_one_memb01_and_exp Then page_to_review = "1"
				If page_display = show_pg_one_address 	Then page_to_review = "2"
				If page_display = show_pg_memb_list 	Then page_to_review = "3"
				If page_display = show_case_info 		Then page_to_review = "4"
				If page_display = show_qual 			Then page_to_review = "5"
				If page_display = show_pg_last			Then page_to_review = "6"
				current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing =  trim(current_listing)
				' MsgBox "Page to Review - " & page_to_review & vbCr & "Current Listing - " & current_listing
				If current_listing = page_to_review Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
					If current_listing = "1"  Then tagline = ": INTVW / CAF 1"        'Adding a specific tagline to the header for the errors
					If current_listing = "2"  Then tagline = ": CAF ADDR"
					If current_listing = "3"  Then tagline = ": CAF MEMBs"
					If current_listing = "4"  Then tagline = ": DETAILS"
					If current_listing = "5" Then tagline = ": CAF QUAL Q"
					If current_listing = "6" Then tagline = ": CAF Last Page"
					If error_message = "" Then error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
					message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

					error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
				End If
			End If
        Next
		If error_message = "" then the_err_msg = ""
		' MsgBox error_message
        'This is the display of all of the messages.
		show_msg = False
        If show_err_msg_during_movement = True Then show_msg = True
		If page_display = show_pg_last AND ButtonPressed <> finish_interview_btn Then show_msg = False

		If ButtonPressed = exp_income_guidance_btn Then show_msg = False
		If ButtonPressed = verif_button Then show_msg = False
		If ButtonPressed = open_hsr_manual_transfer_page_btn Then show_msg = False
		If ButtonPressed >= 500 AND ButtonPressed < 1200 Then show_msg = False
		If ButtonPressed >= 4000 Then show_msg = False

		If error_message = "" Then show_msg = False
		If ButtonPressed = finish_interview_btn Then show_msg = True
		If expedited_determination_needed = True Then
			If expedited_determination_completed = True AND page_display = show_pg_last Then
				If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
			End If
		ElseIf page_display = show_pg_last Then
			If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
		End If

		If show_msg = True Then view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")
		If show_msg = False then the_err_msg = ""
        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE AND show_err_msg_during_movement = False Then
            If back_to_dialog = "1"  Then ButtonPressed = caf_page_one_btn         'This calls another function to go to the first dialog that had an error
            If back_to_dialog = "2"  Then ButtonPressed = caf_addr_btn
            If back_to_dialog = "3"  Then ButtonPressed = caf_membs_btn
            If back_to_dialog = "4"  Then ButtonPressed = case_details_btn
            If back_to_dialog = "5" Then ButtonPressed = caf_qual_q_btn
            If back_to_dialog = "6" Then ButtonPressed = caf_last_page_btn
			If back_to_dialog = "7" Then ButtonPressed = expedited_determination_btn

            Call dialog_movement          'this is where the navigation happens
        End If
    End If
End Function

function display_expedited_dialog()
	expedited_determination_completed = True

	next_btn = 2
	finish_btn = 3

	amounts_btn 		= 10
	determination_btn 	= 20
	review_btn 			= 30

	income_calc_btn								= 100
	asset_calc_btn								= 110
	housing_calc_btn							= 120
	utility_calc_btn							= 130
	snap_active_in_another_state_btn			= 140
	case_previously_had_postponed_verifs_btn	= 150
	household_in_a_facility_btn					= 160

	knowledge_now_support_btn		= 500
	te_02_10_01_btn					= 510

	hsr_manual_expedited_snap_btn 	= 1000
	hsr_applications_btn			= 1100
	ryb_exp_identity_btn			= 1200
	ryb_exp_timeliness_btn			= 1300
	sir_exp_flowchart_btn			= 1400
	cm_04_04_btn					= 1500
	cm_04_06_btn					= 1600
	ht_id_in_solq_btn				= 1700
	cm_04_12_btn					= 1800
	ebt_card_info_btn 	= 1900

	exp_page_display = show_exp_pg_amounts

	Do
		err_msg = ""
		If exp_page_display = show_exp_pg_determination Then Call determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
		If exp_page_display = show_exp_pg_review Then Call determine_actions(case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, is_elig_XFS, snap_denial_date, approval_date, CAF_datestamp, do_we_have_applicant_id, action_due_to_out_of_state_benefits, mn_elig_begin_date, other_snap_state, case_has_previously_postponed_verifs_that_prevent_exp_snap, delay_action_due_to_faci, deny_snap_due_to_faci)

		If determined_income = "" Then determined_income = 0
		If determined_assets = "" Then determined_assets = 0
		If determined_shel = "" Then determined_shel = 0
		If determined_utilities = "" Then determined_utilities = 0
		If calculated_resources = "" Then calculated_resources = 0
		If calculated_expenses = "" Then calculated_expenses = 0
		determined_income = FormatNumber(determined_income, 2, -1, 0, -1) & ""
		determined_assets = FormatNumber(determined_assets, 2, -1, 0, -1) & ""
		determined_shel = FormatNumber(determined_shel, 2, -1, 0, -1) & ""
		determined_utilities = FormatNumber(determined_utilities, 2, -1, 0, -1)
		calculated_resources = FormatNumber(calculated_resources, 2, -1, 0, -1)
		calculated_expenses = FormatNumber(calculated_expenses, 2, -1, 0, -1)

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 555, 385, "Full Expedited Determination"
		  ButtonGroup ButtonPressed
			If exp_page_display = show_exp_pg_amounts then
				Text 504, 12, 65, 10, "Amounts"

				GroupBox 5, 5, 390, 75, "Expedited Screening"
				' If exp_screening_note_found = True Then
				Text 10, 20, 145, 10, "Information pulled from previous case note."
				Text 20, 35, 70, 10, "Income from CAF1: $ "
				Text 100, 35, 80, 10, exp_q_1_income_this_month
				Text 195, 35, 65, 10, "Assets from CAF1: $ "
				Text 270, 35, 75, 10, exp_q_2_assets_this_month
				Text 20, 50, 90, 10, "Housing from CAF1: $ "
				Text 100, 50, 65, 10, exp_q_3_rent_this_month
				Text 195, 50, 65, 10, "Utilities from CAF1: $ "
				Text 270, 50, 75, 10, exp_q_4_utilities_this_month
				Text 15, 65, 160, 10, expedited_screening

				Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
				GroupBox 5, 105, 390, 110, "Information about Income, Resources, and Expenses"
				Text 15, 125, 60, 10, "Gross Income:    $"
				EditBox 75, 120, 155, 15, determined_income
				Text 15, 145, 35, 10, "Assets:   $"
				EditBox 50, 140, 180, 15, determined_assets
				Text 15, 165, 70, 10, "Shelter Expense:    $"
				EditBox 85, 160, 145, 15, determined_shel
				Text 15, 185, 60, 10, "Utilities Expense:"
				Text 77, 185, 145, 15, "$  " & determined_utilities
				PushButton 255, 120, 120, 13, "Calculate Income", income_calc_btn
				PushButton 255, 140, 120, 13, "Calculate Assets", asset_calc_btn
				PushButton 255, 160, 120, 13, "Calculate Housing Cost", housing_calc_btn
				PushButton 255, 180, 120, 13, "Calculate Utilities", utility_calc_btn
				Text 15, 200, 250, 10, "Blank amounts will be defaulted to ZERO."
			End If

			If exp_page_display = show_exp_pg_determination then
				Text 495, 27, 65, 10, "Determination"

				If is_elig_XFS = True Then Text 0, 25, 400, 10, "---------------------------------------------- This case IS EXPEDITED based on this critera: "
				If is_elig_XFS = False Then Text 0, 25, 400, 10, "---------------------------------------------- This case is NOT expedited based on this critera: "

				GroupBox 5, 5, 470, 135, "Expedited Determination"
				Text 15, 50, 120, 10, "Determination Amounts Entered:"
				Text 130, 50, 85, 10, "Total App Month Income:"
				Text 220, 50, 40, 10, "$ " & determined_income
				Text 130, 60, 85, 10, "Total App Month Assets:"
				Text 220, 60, 40, 10, "$ " & determined_assets
				Text 130, 70, 85, 10, "Total App Month Housing:"
				Text 220, 70, 40, 10, "$ " & determined_shel
				Text 130, 80, 85, 10, "Total App Month Utility:"
				Text 220, 80, 40, 10, "$ " & determined_utilities
				Text 295, 50, 135, 10, "Combined Resources (Income + Assets):"
				Text 430, 50, 40, 10, "$ " & calculated_resources
				Text 330, 70, 100, 10, "Combined Housing Expense:"
				Text 430, 70, 40, 10, "$ " & calculated_expenses

				GroupBox 5, 15, 470, 25, ""

				Text 295, 95, 125, 20, "Unit has less than $150 monthly Gross Income AND $100 or less in assets:"
				Text 430, 100, 35, 10, calculated_low_income_asset_test
				Text 295, 115, 125, 20, "Unit's combined resources are less than housing expense:"
				Text 430, 120, 35, 10, calculated_resources_less_than_expenses_test

				Text 18, 90, 65, 10, "Date of Application:"
				Text 85, 90, 50, 10, CAF_datestamp
				Text 25, 100, 60, 10, "Date of Interview:"
				Text 85, 100, 50, 10, interview_date
				Text 25, 115, 60, 10, "Date of Approval:"
				EditBox 85, 110, 60, 15, approval_date
				Text 85, 125, 75, 10, "(or planned approval)"

				GroupBox 5, 135, 470, 155, "Possible Approval Delays"
				Text 95, 150, 205, 10, "Is there a document for proof of identity of the applicant on file?"
				DropListBox 300, 145, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_on_file_yn
				Text 95, 165, 200, 10, "Can the Identity of the applicant be cleard through SOLQ/SMI?"
				DropListBox 300, 160, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", applicant_id_through_SOLQ
				PushButton 350, 160, 120, 13, "HOT TOPIC - Using SOLQ for ID", ht_id_in_solq_btn
				Text 10, 185, 85, 10, "Explain Approval Delays:"
				EditBox 95, 180, 375, 15, delay_explanation
				Text 175, 205, 80, 10, "Specifc case situations:"
				PushButton 255, 200, 215, 15, "SNAP is Active in Another State in " & MAXIS_footer_month & "/" & MAXIS_footer_year, snap_active_in_another_state_btn
				PushButton 255, 215, 215, 15, "Expedited Approved Previously with Postponed Verifications", case_previously_had_postponed_verifs_btn
				PushButton 255, 230, 215, 15, "Household is Currently in a Facility", household_in_a_facility_btn
				Text 15, 255, 330, 10, "If it is already determined that SNAP should be denied, enter a denial date and explanation of denial."
				Text 355, 255, 65, 10, "SNAP Denial Date:"
				EditBox 420, 250, 50, 15, snap_denial_date
				Text 30, 275, 65, 10, "Denial Explanation:"
				EditBox 95, 270, 375, 15, snap_denial_explain
			End If
			If exp_page_display = show_exp_pg_review then
				Text 507, 42, 65, 10, "Review"

				GroupBox 5, 5, 470, 115, "Actions to Take"
				Text 20, 30, 45, 10, "Next Steps:"

				Text 15, 20, 280, 10, case_assesment_text

				Text 25, 40, 435, 20, next_steps_one
				Text 25, 60, 435, 20, next_steps_two
				Text 25, 80, 435, 20, next_steps_three
				Text 25, 100, 435, 20, next_steps_four

				EditBox 800, 800, 50, 15, fake_box_that_does_nothing
				Text 310, 15, 100, 10, "For help with the next steps:"
				PushButton 310, 25, 155, 13, "Request Support from Knowledge Now", knowledge_now_support_btn

				GroupBox 5, 120, 470, 85, "Postponed Verifications"
				If is_elig_XFS = True AND IsDate(snap_denial_date) = False Then
					Text 15, 135, 160, 10, "Are there Postponed Verifications for this case?"
					DropListBox 180, 130, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", postponed_verifs_yn
					Text 20, 155, 80, 10, "Postponed Verifications:"
					EditBox 105, 150, 360, 15, list_postponed_verifs
					PushButton 320, 130, 145, 13, "TE 02.10.01 EXP w/ Pending Verifs", te_02_10_01_btn
					Text 20, 175, 120, 10, "Can I postpone Verifications for ..."
					Text 145, 175, 70, 10, "Immigration - YES."
					Text 225, 175, 55, 10, "Sponsor - YES."
					Text 300, 175, 125, 10, "anything OTHER than ID - YES. "
					Text 30, 190, 300, 10, "Applicant's identity is the ONLY required verification to approve Expedited SNAP."
					PushButton 320, 187, 145, 13, "CM 04.12 Verification Requirement for EXP", cm_04_12_btn
				End If
				If is_elig_XFS = False Then
					Text 15, 135, 450, 10, "We cannot postpone any verifications for a case that does not meet Expedited criteria."
				End If
				If IsDate(snap_denial_date) = True Then
					Text 15, 135, 450, 10, "Additional verifications are not needed if a Denial has already been determined."
				End If

				GroupBox 5, 205, 470, 70, "EBT Information"
				If IsDate(snap_denial_date) = True Then
					Text 15, 220, 415, 10, "Advise resident to keep track of an EBT card they have received, even though the application is being denied."
					Text 20, 235, 415, 10, "If the case ever reapplies, or is determined eligible, the EBT card remains connected to the case and getting benefits will be easier."
				Else
					Text 15, 220, 335, 10, "Do not delay in approving SNAP benefits due to if the household does or does not have an EBT card."
					Text 20, 235, 415, 10, "If there has never been a card issued for a case, approving the benefit with an REI will prevent a card from being sent via mail."
					Text 20, 245, 305, 10, "If a case needs the first card mailed, do NOT REI benefits as they will not receive their card."
				End If
				Text 15, 260, 255, 10, "EBT Card issues can be complicated. Refer to the EBT Card Information here:"
				PushButton 270, 257, 195, 13, "Information about EBT Cards", ebt_card_info_btn

			End If
			GroupBox 5, 315, 470, 60, "If you need support in handling for expedited, please access these resources:"
			PushButton 15, 325, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
			PushButton 15, 340, 150, 13, "HSR Manual - Applications", hsr_applications_btn
			PushButton 15, 355, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			PushButton 165, 325, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
			PushButton 165, 340, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			PushButton 315, 325, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			PushButton 315, 340, 150, 13, "CM 04.06 - 1st Month Processing", cm_04_06_btn

			If exp_page_display <> show_exp_pg_amounts then PushButton 485, 10, 65, 13, "Amounts", amounts_btn
			If exp_page_display <> show_exp_pg_determination then PushButton 485, 25, 65, 13, "Determination", determination_btn
			If exp_page_display <> show_exp_pg_review then PushButton 485, 40, 65, 13, "Review", review_btn
			If exp_page_display <> show_exp_pg_review then PushButton 500, 365, 50, 15, "Next", next_btn
			If exp_page_display = show_exp_pg_review then PushButton 500, 365, 50, 15, "Return", finish_btn
		EndDialog

		Dialog Dialog1

		If ButtonPressed = -1 Then
			If exp_page_display <> show_exp_pg_review then ButtonPressed = next_btn
			If exp_page_display = show_exp_pg_review then ButtonPressed = finish_btn
		End If

		If ButtonPressed = income_calc_btn Then Call app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, EXP_JOBS_ARRAY, EXP_BUSI_ARRAY, EXP_UNEA_ARRAY)
		If ButtonPressed = asset_calc_btn Then Call app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, EXP_ACCT_ARRAY)
		If ButtonPressed = housing_calc_btn Then Call app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
		If ButtonPressed = utility_calc_btn Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
		If ButtonPressed = snap_active_in_another_state_btn Then
			If IsDate(CAF_datestamp) = False Then MsgBox "Attention:" & vbCr & vbCr & "The funcationality to determine actions if a household is reporting benefits in another state cannot be run if a valid application date has not been entered."
			If IsDate(CAF_datestamp) = True Then Call snap_in_another_state_detail(CAF_datestamp, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
		End If
		If ButtonPressed = case_previously_had_postponed_verifs_btn Then Call previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_CAF_datestamp, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
		If ButtonPressed = household_in_a_facility_btn Then Call household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)

		If ButtonPressed = knowledge_now_support_btn Then
			Call send_support_email_to_KN
			STATS_manualtime = STATS_manualtime + 300
		End If
		If ButtonPressed = te_02_10_01_btn Then Call view_poli_temp("02", "10", "01", "")

		If exp_page_display = show_exp_pg_determination Then
			delay_due_to_interview = False
			do_we_have_applicant_id = "UNKNOWN"
			If applicant_id_on_file_yn = "Yes" OR applicant_id_through_SOLQ = "Yes" Then do_we_have_applicant_id = True
			If applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "No" Then do_we_have_applicant_id = False

			If IsDate(snap_denial_date) = True Then
				If DateDiff("d", date, snap_denial_date) > 0 Then err_msg = err_msg & vbCr & "* Future Date denials or 'Possible' denials are not what the 'SNAP Denial Date' field is for." & vbCr &_
																						  "* Only indicate a denial if you already have enough information to determine that the SNAP application should be denied." & vbCr &_
																						  "* If this is the determination, review the date in the SNAP Denial Field as it appears to be a future date."
				snap_denial_explain = trim(snap_denial_explain)
				If len(snap_denial_explain) < 10 then err_msg = err_msg & vbCr & "* Since this SNAP case is to be denied, explain the reason for denial in detail."
			Else
				If is_elig_XFS = True Then
					If IsDate(approval_date) = True Then
						If DateDiff("d", date, approval_date) > 0 Then err_msg = err_msg & vbCr & "* Approvals should happen the same day an Expedited Determination is completed if the case is Expedited. Since the Income, Assets, and Expenses indicate this case is expedited AND we appear to be ready to approve, this should be completed today."
						' If DateDiff("d", interview_date, date) < 0 Then
					End If
					If applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Indicate if we have identity of the applicant on file or available through SOLQ"
					ElseIf applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "?" Then
						err_msg = err_msg & vbCr & "* Since there is no identity found in the file for the applicant, check SOLQ/SMI to verify identity."
					ElseIf applicant_id_on_file_yn = "?" AND applicant_id_through_SOLQ = "No" Then
						err_msg = err_msg & vbCr & "* Since the applicant's identity cannot be cleared through SOLQ/SMI, check the case file and person file for documents that can be used to verify identity. Remember that SNAP does NOT require a Photo ID or Official Government ID."
					End If

					'Defaulting Delay Explanation
					If IsDate(approval_date) = True AND IsDate(interview_date) = True AND IsDate(CAF_datestamp) = True Then
						If DateDiff("d", CAF_datestamp, approval_date) > 7 Then
							If DateDiff("d", interview_date, approval_date) = 0 Then delay_due_to_interview = True
						End If
					End If
					If delay_due_to_interview = True AND InStr(delay_explanation, "Approval of Expedited delayed until completion of Interview") = 0 Then
						delay_explanation = delay_explanation & "; Approval of Expedited delayed until completion of Interview."
					End If
					If delay_due_to_interview = False then
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview.", "")
						delay_explanation = replace(delay_explanation, "Approval of Expedited delayed until completion of Interview", "")
					End If
					If do_we_have_applicant_id = False AND InStr(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant") = 0 Then
						delay_explanation = delay_explanation & "; Approval cannot be completed as we have NO Proof of Identity for the Applicant."
					End If
					If do_we_have_applicant_id <> False Then
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant.", "")
						delay_explanation = replace(delay_explanation, "Approval cannot be completed as we have NO Proof of Identity for the Applicant", "")
					End If

					Call format_explanation_text(delay_explanation)
					Call format_explanation_text(snap_denial_explain)

					expedited_approval_delayed = False
					If IsDate(approval_date) = False Then expedited_approval_delayed = True
					If IsDate(approval_date) = True  AND IsDate(CAF_datestamp) = True Then
						If DateDiff("d", CAF_datestamp, approval_date) > 7 Then expedited_approval_delayed = True
					End If
					If expedited_approval_delayed = True AND len(delay_explanation) < 20 Then err_msg = err_msg & vbCR & "* The approval of the Expedited SNAP is or has been delayed. Provide a detailed explaination of the reason for delay or complete the approval."

				End If
				If is_elig_XFS = False Then

				End If
			End If

		End If
		If exp_page_display = show_exp_pg_review Then
			If postponed_verifs_yn = "Yes" AND trim(list_postponed_verifs) = "" Then err_msg = err_msg & vbCr & "* Since you have Postponed Verifications indicated, list what they are for the NOTE."
		End If

		If ButtonPressed = next_btn AND err_msg = "" Then exp_page_display = exp_page_display + 1
		If ButtonPressed = amounts_btn Then exp_page_display = show_exp_pg_amounts
		If ButtonPressed = determination_btn AND err_msg = "" Then exp_page_display = show_exp_pg_determination
		If ButtonPressed = review_btn AND err_msg = "" AND exp_page_display <> show_exp_pg_amounts Then exp_page_display = show_exp_pg_review
		If ButtonPressed = review_btn AND err_msg = "" AND exp_page_display = show_exp_pg_amounts Then exp_page_display = show_exp_pg_determination

		If ButtonPressed = 0 then
			err_msg = ""
			expedited_determination_completed = False
		End If
		If err_msg <> "" And ButtonPressed < 100 AND exp_page_display <> show_exp_pg_amounts Then MsgBox "***** Action Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & err_msg

		If ButtonPressed <> finish_btn Then err_msg = "LOOP"
		If ButtonPressed = 0 then err_msg = ""

		If ButtonPressed >= 1000 Then
			If ButtonPressed = hsr_manual_expedited_snap_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Expedited_SNAP.aspx"
			If ButtonPressed = hsr_applications_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Applications.aspx"
			If ButtonPressed = ryb_exp_identity_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%201%20-%20Identity.mp4"
			If ButtonPressed = ryb_exp_timeliness_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/Retrain_Your_Brain/SNAP%20Expedited%202%20-%20Timeliness.mp4"
			If ButtonPressed = sir_exp_flowchart_btn Then resource_URL = "https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP%20Expedited%20Service%20Flowchart.pdf"
			If ButtonPressed = cm_04_04_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000404"
			If ButtonPressed = cm_04_06_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000406"
			If ButtonPressed = ht_id_in_solq_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/SitePages/How-to-use-SMI-SOLQ-to-verify-ID-for-SNAP.aspx"
			If ButtonPressed = cm_04_12_btn Then resource_URL = "https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_000412"
			If ButtonPressed = ebt_card_info_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Accounting.aspx#%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8B%E2%80%8Bprocesses-for-receiving-ebt-cards-at-the-county-offices"

			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe " & resource_URL
		End If

	Loop until err_msg = ""

	page_display = show_pg_last
end function

function evaluate_for_expedited(app_month_income, app_month_assets, app_month_housing_cost, heat_checkbox, air_checkbox, electric_checkbox, phone_checkbox, app_month_utilities_cost, app_month_expenses, case_is_expedited)
	If heat_checkbox = checked OR air_checkbox = checked Then
        app_month_utilities_cost = heat_AC_amt
	ElseIf electric_checkbox = checked AND phone_checkbox = checked Then
		app_month_utilities_cost = electric_amt + phone_amt
	ElseIf electric_checkbox = checked Then
		app_month_utilities_cost = electric_amt
	ElseIf phone_checkbox = checked Then
		app_month_utilities_cost = phone_amt
	End If
	If app_month_housing_cost = "" Then app_month_housing_cost = 0
	app_month_housing_cost = app_month_housing_cost * 1
	app_month_expenses = app_month_utilities_cost + app_month_housing_cost

	If app_month_income = "" Then app_month_income = 0
	app_month_income = app_month_income * 1

	If app_month_assets = "" Then app_month_assets = 0
	app_month_assets = app_month_assets * 1

	income_and_assets = app_month_income + app_month_assets

	case_is_expedited = False
	If app_month_income < 150 AND app_month_assets <= 100 Then case_is_expedited = True
	If income_and_assets < app_month_expenses Then case_is_expedited = True
	app_month_income = app_month_income & ""
	app_month_assets = app_month_assets & ""
	app_month_housing_cost = app_month_housing_cost & ""
end function

'This function records the variables into a txt file so that it can be retrieved by the script if run later.
function save_your_work()

	'Now determines name of file
	If MAXIS_case_number <> "" Then
		save_your_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
	End If

	With (CreateObject("Scripting.FileSystemObject"))

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(save_your_work_path) = True then
			.DeleteFile(save_your_work_path)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

		If .FileExists(save_your_work_path) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(save_your_work_path, ForWriting, true)

			'Write the contents of the text file
			If IsNumeric(add_to_time) = True Then objTextStream.WriteLine "TIME SPENT - "	& timer - start_time + add_to_time
			If IsNumeric(add_to_time) = False Then objTextStream.WriteLine "TIME SPENT - "	& timer - start_time

			objTextStream.WriteLine "CAF - DATE - " & CAF_datestamp

            objTextStream.WriteLine "MFIP - ORNT - " & MFIP_orientation_assessed_and_completed
            objTextStream.WriteLine "MFIP - DWP - " & family_cash_program
            objTextStream.WriteLine "FMCA - 01 - " & famliy_cash_notes

			objTextStream.WriteLine "PROG - CASH - " & cash_other_req_detail
			objTextStream.WriteLine "PROG - SNAP - " & snap_other_req_detail
			objTextStream.WriteLine "PROG - EMER - " & emer_other_req_detail
			If CASH_on_CAF_checkbox = checked Then objTextStream.WriteLine "CASH PROG CHECKED"
			If SNAP_on_CAF_checkbox = checked Then objTextStream.WriteLine "SNAP PROG CHECKED"
			If EMER_on_CAF_checkbox = checked Then objTextStream.WriteLine "EMER PROG CHECKED"

			objTextStream.WriteLine "CASH - TYPE - " & type_of_cash
			objTextStream.WriteLine "PROC - CASH - " & the_process_for_cash
			objTextStream.WriteLine "CASH - RVMO - " & next_cash_revw_mo
			objTextStream.WriteLine "CASH - RVYR - " & next_cash_revw_yr

			objTextStream.WriteLine "PROC - SNAP - " & the_process_for_snap
			objTextStream.WriteLine "SNAP - RVMO - " & next_snap_revw_mo
			objTextStream.WriteLine "SNAP - RVYR - " & next_snap_revw_yr

			objTextStream.WriteLine "EMER - TYPE - " & type_of_emer
			objTextStream.WriteLine "PROC - EMER - " & the_process_for_emer

			objTextStream.WriteLine "PRE - ATC - " & all_the_clients
			objTextStream.WriteLine "PRE - WHO - " & who_are_we_completing_the_interview_with
			objTextStream.WriteLine "PRE - HOW - " & how_are_we_completing_the_interview
			objTextStream.WriteLine "PRE - ITP - " & interpreter_information
			objTextStream.WriteLine "PRE - LNG - " & interpreter_language
			objTextStream.WriteLine "PRE - AID - " & arep_interview_id_information
			objTextStream.WriteLine "PRE - DET - " & non_applicant_interview_info

			objTextStream.WriteLine "EXP - 1 - " & exp_q_1_income_this_month
			objTextStream.WriteLine "EXP - 2 - " & exp_q_2_assets_this_month
			objTextStream.WriteLine "EXP - 3 - RENT - " & exp_q_3_rent_this_month
			If caf_exp_pay_heat_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - HEAT"
			If caf_exp_pay_ac_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - ACON"
			If caf_exp_pay_electricity_checkbox = checked 	Then objTextStream.WriteLine "EXP - 3 - ELEC"
			If caf_exp_pay_phone_checkbox = checked 		Then objTextStream.WriteLine "EXP - 3 - PHON"
			If caf_exp_pay_none_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - NONE"
			objTextStream.WriteLine "EXP - 3 - UTIL - " & exp_q_4_utilities_this_month
			objTextStream.WriteLine "EXP - 4 - " & exp_migrant_seasonal_formworker_yn
			objTextStream.WriteLine "EXP - 5 - PREV - " & exp_received_previous_assistance_yn
			objTextStream.WriteLine "EXP - 5 - WHEN - " & exp_previous_assistance_when
			objTextStream.WriteLine "EXP - 5 - WHER - " & exp_previous_assistance_where
			objTextStream.WriteLine "EXP - 5 - WHAT - " & exp_previous_assistance_what
			objTextStream.WriteLine "EXP - 6 - PREG - " & exp_pregnant_yn
			objTextStream.WriteLine "EXP - 6 - WHO? - " & exp_pregnant_who
			objTextStream.WriteLine "EXP - INTVW - INCM - " & intv_app_month_income
			objTextStream.WriteLine "EXP - INTVW - ASST - " & intv_app_month_asset
			objTextStream.WriteLine "EXP - INTVW - RENT - " & intv_app_month_housing_expense
			If intv_exp_pay_heat_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - HEAT"
			If intv_exp_pay_ac_checkbox = checked 			Then objTextStream.WriteLine "EXP - INTVW - ACON"
			If intv_exp_pay_electricity_checkbox = checked 	Then objTextStream.WriteLine "EXP - INTVW - ELEC"
			If intv_exp_pay_phone_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - PHON"
			If intv_exp_pay_none_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - NONE"
			objTextStream.WriteLine "EXP - INTVW - ID - " & id_verif_on_file
			objTextStream.WriteLine "EXP - INTVW - 89 - " & snap_active_in_other_state
			objTextStream.WriteLine "EXP - INTVW - EXP - " & last_snap_was_exp

			objTextStream.WriteLine "ADR - RESI - STR - " & resi_addr_street_full
			objTextStream.WriteLine "ADR - RESI - CIT - " & resi_addr_city
			objTextStream.WriteLine "ADR - RESI - STA - " & resi_addr_state
			objTextStream.WriteLine "ADR - RESI - ZIP - " & resi_addr_zip

			objTextStream.WriteLine "ADR - RESI - RES - " & reservation_yn
			objTextStream.WriteLine "ADR - RESI - NAM - " & reservation_name

			objTextStream.WriteLine "ADR - RESI - HML - " & homeless_yn

			objTextStream.WriteLine "ADR - RESI - LIV - " & living_situation

			objTextStream.WriteLine "ADR - MAIL - STR - " & mail_addr_street_full
			objTextStream.WriteLine "ADR - MAIL - CIT - " & mail_addr_city
			objTextStream.WriteLine "ADR - MAIL - STA - " & mail_addr_state
			objTextStream.WriteLine "ADR - MAIL - ZIP - " & mail_addr_zip

			objTextStream.WriteLine "ADR - PHON - NON - " & phone_one_number
			objTextStream.WriteLine "ADR - PHON - TON - " & phone_one_type
			objTextStream.WriteLine "ADR - PHON - NTW - " & phone_two_number
			objTextStream.WriteLine "ADR - PHON - TTW - " & phone_two_type
			objTextStream.WriteLine "ADR - PHON - NTH - " & phone_three_number
			objTextStream.WriteLine "ADR - PHON - TTH - " & phone_three_type

			objTextStream.WriteLine "ADR - DATE - " & address_change_date
			objTextStream.WriteLine "ADR - CNTY - " & resi_addr_county

			objTextStream.WriteLine "CASE DETAILS - HH - " & household_info
			objTextStream.WriteLine "CASE DETAILS - EI - " & earned_income_info
			objTextStream.WriteLine "CASE DETAILS - UNEA - " & unearned_income_info
			objTextStream.WriteLine "CASE DETAILS - SHEL - " & housing_expanese_info
			objTextStream.WriteLine "CASE DETAILS - EXP - " & other_expenses_info
			objTextStream.WriteLine "CASE DETAILS - AST - " & assets_info
			objTextStream.WriteLine "CASE DETAILS - OTHR - " & other_notes_info
			objTextStream.WriteLine "CASE DETAILS - CHG - " & form_changes_info

			If rights_responsibilities_checkbox = checked Then objTextStream.WriteLine "rights_responsibilities_checkbox"
			If privacy_practices_checkbox = checked Then objTextStream.WriteLine "privacy_practices_checkbox"
			If ebt_card_checkbox = checked Then objTextStream.WriteLine "ebt_card_checkbox"
			If complaints_civil_rights_checkbox = checked Then objTextStream.WriteLine "complaints_civil_rights_checkbox"
			If reporting_resp_checkbox = checked Then objTextStream.WriteLine "reporting_resp_checkbox"
			If program_info_checkbox = checked Then objTextStream.WriteLine "program_info_checkbox"
			If child_support_info_checkbox = checked Then objTextStream.WriteLine "child_support_info_checkbox"
			If mfip_minor_caregiver_checkbox = checked Then objTextStream.WriteLine "mfip_minor_caregiver_checkbox"
			If renewal_info_checkbox = checked Then objTextStream.WriteLine "renewal_info_checkbox"
			If ievs_info_checkbox = checked Then objTextStream.WriteLine "ievs_info_checkbox"
			If appeal_rights_checkbox = checked Then objTextStream.WriteLine "appeal_rights_checkbox"

			objTextStream.WriteLine "QQ1A - " & qual_question_one
			objTextStream.WriteLine "QQ1M - " & qual_memb_one
			objTextStream.WriteLine "QQ2A - " & qual_question_two
			objTextStream.WriteLine "QQ2M - " & qual_memb_two
			objTextStream.WriteLine "QQ3A - " & qual_question_three
			objTextStream.WriteLine "QQ3M - " & qual_memb_three
			objTextStream.WriteLine "QQ4A - " & qual_question_four
			objTextStream.WriteLine "QQ4M - " & qual_memb_four
			objTextStream.WriteLine "QQ5A - " & qual_question_five
			objTextStream.WriteLine "QQ5M - " & qual_memb_five

			objTextStream.WriteLine "AREP - 001 - " & arep_in_MAXIS
			objTextStream.WriteLine "AREP - 002 - " & MAXIS_arep_updated
			objTextStream.WriteLine "AREP - 003 - " & arep_authorization
			objTextStream.WriteLine "AREP - 004 - " & arep_authorized

			objTextStream.WriteLine "AREP - 01 - " & arep_name
			objTextStream.WriteLine "AREP - 02 - " & arep_relationship
			objTextStream.WriteLine "AREP - 03 - " & arep_phone_number
			objTextStream.WriteLine "AREP - 04 - " & arep_addr_street
			objTextStream.WriteLine "AREP - 05 - " & arep_addr_city
			objTextStream.WriteLine "AREP - 06 - " & arep_addr_state
			objTextStream.WriteLine "AREP - 07 - " & arep_addr_zip
			If arep_complete_forms_checkbox = checked Then objTextStream.WriteLine "AREP - 08"
			If arep_get_notices_checkbox = checked Then objTextStream.WriteLine "AREP - 09"
			If arep_use_SNAP_checkbox = checked Then objTextStream.WriteLine "AREP - 10"
			If arep_on_CAF_checkbox = checked Then objTextStream.WriteLine "AREP - 11"
			objTextStream.WriteLine "AREP - 12 - " & arep_action

			objTextStream.WriteLine "MX-AREP - 01 - " & MAXIS_arep_name
			objTextStream.WriteLine "MX-AREP - 02 - " & MAXIS_arep_relationship
			objTextStream.WriteLine "MX-AREP - 03 - " & MAXIS_arep_phone_number
			objTextStream.WriteLine "MX-AREP - 04 - " & MAXIS_arep_addr_street
			objTextStream.WriteLine "MX-AREP - 05 - " & MAXIS_arep_addr_city
			objTextStream.WriteLine "MX-AREP - 06 - " & MAXIS_arep_addr_state
			objTextStream.WriteLine "MX-AREP - 07 - " & MAXIS_arep_addr_zip

			objTextStream.WriteLine "CAF-AREP - 01 - " & CAF_arep_name
			objTextStream.WriteLine "CAF-AREP - 02 - " & CAF_arep_relationship
			objTextStream.WriteLine "CAF-AREP - 03 - " & CAF_arep_phone_number
			objTextStream.WriteLine "CAF-AREP - 04 - " & CAF_arep_addr_street
			objTextStream.WriteLine "CAF-AREP - 05 - " & CAF_arep_addr_city
			objTextStream.WriteLine "CAF-AREP - 06 - " & CAF_arep_addr_state
			objTextStream.WriteLine "CAF-AREP - 07 - " & CAF_arep_addr_zip
			If CAF_arep_complete_forms_checkbox = checked Then objTextStream.WriteLine "CAF-AREP - 08"
			If CAF_arep_get_notices_checkbox = checked Then objTextStream.WriteLine "CAF-AREP - 09"
			If CAF_arep_use_SNAP_checkbox = checked Then objTextStream.WriteLine "CAF-AREP - 10"
			objTextStream.WriteLine "CAF-AREP - 11 - " & CAF_arep_action

			objTextStream.WriteLine "SIG - 01 - " & signature_detail
			objTextStream.WriteLine "SIG - 02 - " & signature_person
			objTextStream.WriteLine "SIG - 03 - " & signature_date
			objTextStream.WriteLine "SIG - 04 - " & second_signature_detail
			objTextStream.WriteLine "SIG - 05 - " & second_signature_person
			objTextStream.WriteLine "SIG - 06 - " & second_signature_date
			objTextStream.WriteLine "SIG - 07 - " & client_signed_verbally_yn
			objTextStream.WriteLine "SIG - 08 - " & interview_date
			objTextStream.WriteLine "ASSESS - 01 - " & exp_snap_approval_date
			objTextStream.WriteLine "ASSESS - 02 - " & exp_snap_delays
			objTextStream.WriteLine "ASSESS - 03 - " & snap_denial_date
			objTextStream.WriteLine "ASSESS - 04 - " & snap_denial_explain
			objTextStream.WriteLine "ASSESS - 05 - " & pend_snap_on_case

			objTextStream.WriteLine "ASSESS - 06 - " & family_cash_case_yn
			objTextStream.WriteLine "ASSESS - 07 - " & absent_parent_yn
			objTextStream.WriteLine "ASSESS - 08 - " & relative_caregiver_yn
			objTextStream.WriteLine "ASSESS - 09 - " & minor_caregiver_yn

			objTextStream.WriteLine "CLAR - TOTAL - " & discrepancies_exist
			objTextStream.WriteLine "CLAR - PHONE - 01 - " & disc_no_phone_number
			objTextStream.WriteLine "CLAR - PHONE - 02 - " & disc_phone_confirmation
			objTextStream.WriteLine "CLAR - PHEXP - 01 - " & disc_yes_phone_no_expense
			objTextStream.WriteLine "CLAR - PHEXP - 02 - " & disc_yes_phone_no_expense_confirmation
			objTextStream.WriteLine "CLAR - PHEXP - 03 - " & disc_no_phone_yes_expense
			objTextStream.WriteLine "CLAR - PHEXP - 04 - " & disc_no_phone_yes_expense_confirmation
			objTextStream.WriteLine "CLAR - HOMLS - 01 - " & disc_homeless_no_mail_addr
			objTextStream.WriteLine "CLAR - HOMLS - 02 - " & disc_homeless_confirmation
			objTextStream.WriteLine "CLAR - OTOCO - 01 - " & disc_out_of_county
			objTextStream.WriteLine "CLAR - OTOCO - 02 - " & disc_out_of_county_confirmation
			objTextStream.WriteLine "CLAR - HOUS$ - 01 - " & disc_rent_amounts
			objTextStream.WriteLine "CLAR - HOUS$ - 02 - " & disc_rent_amounts_confirmation
			objTextStream.WriteLine "CLAR - UTIL$ - 01 - " & disc_utility_amounts
			objTextStream.WriteLine "CLAR - UTIL$ - 02 - " & disc_utility_amounts_confirmation

			objTextStream.WriteLine "EXPDET - 01 - " & expedited_determination_completed
			objTextStream.WriteLine "EXPDET - 02 - " & expedited_screening
			objTextStream.WriteLine "EXPDET - 03 - " & calculated_low_income_asset_test
			objTextStream.WriteLine "EXPDET - 04 - " & calculated_resources_less_than_expenses_test
			objTextStream.WriteLine "EXPDET - 05 - " & is_elig_XFS
			objTextStream.WriteLine "EXPDET - 06 - " & case_assesment_text
			objTextStream.WriteLine "EXPDET - 07 - " & next_steps_one
			objTextStream.WriteLine "EXPDET - 08 - " & next_steps_two
			objTextStream.WriteLine "EXPDET - 09 - " & next_steps_three
			objTextStream.WriteLine "EXPDET - 10 - " & next_steps_four
			objTextStream.WriteLine "EXPDET - 11 - " & caf_1_resources
			objTextStream.WriteLine "EXPDET - 12 - " & caf_1_expenses
			objTextStream.WriteLine "EXPDET - 13 - " & applicant_id_on_file_yn
			objTextStream.WriteLine "EXPDET - 14 - " & applicant_id_through_SOLQ
			objTextStream.WriteLine "EXPDET - 15 - " & approval_date
			objTextStream.WriteLine "EXPDET - 16 - " & day_30_from_application
			objTextStream.WriteLine "EXPDET - 17 - " & delay_explanation
			objTextStream.WriteLine "EXPDET - 18 - " & postponed_verifs_yn
			objTextStream.WriteLine "EXPDET - 19 - " & list_postponed_verifs
			objTextStream.WriteLine "EXPDET - 20 - " & first_time_in_exp_det

			objTextStream.WriteLine "EXPDET - 21 - " & income_review_completed
			objTextStream.WriteLine "EXPDET - 22 - " & assets_review_completed
			objTextStream.WriteLine "EXPDET - 23 - " & shel_review_completed
			objTextStream.WriteLine "EXPDET - 24 - " & note_calculation_detail

			objTextStream.WriteLine "EXPDET - INCM - 01 - " & determined_income
			objTextStream.WriteLine "EXPDET - INCM - 02 - " & jobs_income_yn
			objTextStream.WriteLine "EXPDET - INCM - 03 - " & busi_income_yn
			objTextStream.WriteLine "EXPDET - INCM - 04 - " & unea_income_yn
			For each_item = 0 to UBound(EXP_JOBS_ARRAY, 2)
				objTextStream.WriteLine "ARR - EXP_JOBS_ARRAY - " & EXP_JOBS_ARRAY(jobs_employee_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_employer_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_wage_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_hours_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_frequency_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_monthly_pay_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_notes_const, each_item)
			Next
			For each_item = 0 to UBound(EXP_BUSI_ARRAY, 2)
				objTextStream.WriteLine "ARR - EXP_BUSI_ARRAY - " & EXP_BUSI_ARRAY(busi_owner_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_info_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_monthly_earnings_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_annual_earnings_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_notes_const, each_item)
			Next
			For each_item = 0 to UBound(EXP_UNEA_ARRAY, 2)
				objTextStream.WriteLine "ARR - EXP_UNEA_ARRAY - " & EXP_UNEA_ARRAY(unea_owner_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_info_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_monthly_earnings_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_weekly_earnings_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_notes_const, each_item)
			Next

			objTextStream.WriteLine "EXPDET - ASST - 01 - " & determined_assets
			objTextStream.WriteLine "EXPDET - ASST - 02 - " & cash_amount_yn
			objTextStream.WriteLine "EXPDET - ASST - 03 - " & bank_account_yn
			objTextStream.WriteLine "EXPDET - ASST - 04 - " & cash_amount
			For each_item = 0 to UBound(EXP_ACCT_ARRAY, 2)
				objTextStream.WriteLine "ARR - EXP_ACCT_ARRAY - " & EXP_ACCT_ARRAY(account_type_const, each_item)&"~"&EXP_ACCT_ARRAY(account_owner_const, each_item)&"~"&EXP_ACCT_ARRAY(bank_name_const, each_item)&"~"&EXP_ACCT_ARRAY(account_amount_const, each_item)&"~"&EXP_ACCT_ARRAY(account_notes_const, each_item)
			Next

			objTextStream.WriteLine "EXPDET - SHEL - 01 - " & determined_shel
			objTextStream.WriteLine "EXPDET - SHEL - 02 - " & rent_amount
			objTextStream.WriteLine "EXPDET - SHEL - 03 - " & lot_rent_amount
			objTextStream.WriteLine "EXPDET - SHEL - 04 - " & mortgage_amount
			objTextStream.WriteLine "EXPDET - SHEL - 05 - " & insurance_amount
			objTextStream.WriteLine "EXPDET - SHEL - 06 - " & tax_amount
			objTextStream.WriteLine "EXPDET - SHEL - 07 - " & room_amount
			objTextStream.WriteLine "EXPDET - SHEL - 08 - " & garage_amount

			objTextStream.WriteLine "EXPDET - HEST - 01 - " & determined_utilities
			objTextStream.WriteLine "EXPDET - HEST - 02 - " & heat_expense
			objTextStream.WriteLine "EXPDET - HEST - 03 - " & ac_expense
			objTextStream.WriteLine "EXPDET - HEST - 04 - " & electric_expense
			objTextStream.WriteLine "EXPDET - HEST - 05 - " & phone_expense
			objTextStream.WriteLine "EXPDET - HEST - 06 - " & none_expense
			objTextStream.WriteLine "EXPDET - HEST - 07 - " & all_utilities
			objTextStream.WriteLine "EXPDET - RESOURCES - " & calculated_resources
			objTextStream.WriteLine "EXPDET - EXPENSES - " & calculated_expenses


			objTextStream.WriteLine "EXPDET - OUTSTATE - 01 - " & other_snap_state
			objTextStream.WriteLine "EXPDET - OUTSTATE - 02 - " & other_state_reported_benefit_end_date
			objTextStream.WriteLine "EXPDET - OUTSTATE - 03 - " & other_state_benefits_openended
			objTextStream.WriteLine "EXPDET - OUTSTATE - 04 - " & other_state_contact_yn
			objTextStream.WriteLine "EXPDET - OUTSTATE - 05 - " & other_state_verified_benefit_end_date
			objTextStream.WriteLine "EXPDET - OUTSTATE - 06 - " & mn_elig_begin_date
			objTextStream.WriteLine "EXPDET - OUTSTATE - 07 - " & action_due_to_out_of_state_benefits

			objTextStream.WriteLine "EXPDET - PSTPND - 01 - " & case_has_previously_postponed_verifs_that_prevent_exp_snap
			objTextStream.WriteLine "EXPDET - PSTPND - 02 - " & prev_post_verif_assessment_done
			objTextStream.WriteLine "EXPDET - PSTPND - 03 - " & previous_CAF_datestamp
			objTextStream.WriteLine "EXPDET - PSTPND - 04 - " & previous_expedited_package
			objTextStream.WriteLine "EXPDET - PSTPND - 05 - " & prev_verifs_mandatory_yn
			objTextStream.WriteLine "EXPDET - PSTPND - 06 - " & prev_verif_list
			objTextStream.WriteLine "EXPDET - PSTPND - 07 - " & curr_verifs_postponed_yn
			objTextStream.WriteLine "EXPDET - PSTPND - 08 - " & ongoing_snap_approved_yn
			objTextStream.WriteLine "EXPDET - PSTPND - 09 - " & prev_post_verifs_recvd_yn

			objTextStream.WriteLine "EXPDET - FACI - 01 - " & delay_action_due_to_faci
			objTextStream.WriteLine "EXPDET - FACI - 02 - " & deny_snap_due_to_faci
			objTextStream.WriteLine "EXPDET - FACI - 03 - " & faci_review_completed
			objTextStream.WriteLine "EXPDET - FACI - 04 - " & facility_name
			objTextStream.WriteLine "EXPDET - FACI - 05 - " & snap_inelig_faci_yn
			objTextStream.WriteLine "EXPDET - FACI - 06 - " & faci_entry_date
			objTextStream.WriteLine "EXPDET - FACI - 07 - " & faci_release_date
			If release_date_unknown_checkbox = checked Then objTextStream.WriteLine "EXPDET - FACI - 08"
			objTextStream.WriteLine "EXPDET - FACI - 09 - " & release_within_30_days_yn

			objTextStream.WriteLine "VERIFS - " & verifs_selected
			objTextStream.WriteLine "VRFDTE - " & verif_req_form_sent_date
			If number_verifs_checkbox = checked Then objTextStream.WriteLine "NUMBER VERIFS"
			If verifs_postponed_checkbox = checked Then objTextStream.WriteLine "POSTPONE VERIFS"
            If verif_snap_checkbox = checked then objTextStream.WriteLine "verif_snap_checkbox"
            If verif_cash_checkbox = checked then objTextStream.WriteLine "verif_cash_checkbox"
            If verif_mfip_checkbox = checked then objTextStream.WriteLine "verif_mfip_checkbox"
            If verif_dwp_checkbox = checked then objTextStream.WriteLine "verif_dwp_checkbox"
            If verif_msa_checkbox = checked then objTextStream.WriteLine "verif_msa_checkbox"
            If verif_ga_checkbox = checked then objTextStream.WriteLine "verif_ga_checkbox"
            If verif_grh_checkbox = checked then objTextStream.WriteLine "verif_grh_checkbox"
            If verif_emer_checkbox = checked then objTextStream.WriteLine "verif_emer_checkbox"
            If verif_hc_checkbox = checked then objTextStream.WriteLine "verif_hc_checkbox"

			For known_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
				' objTextStream.WriteLine "ARR - ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_last_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_first_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_other_names, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_dob, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_gender, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_former_state, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_citizen, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_written_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_notes, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				race_a_info = ""
				race_b_info = ""
				race_n_info = ""
				race_p_info = ""
				race_w_info = ""
				prog_s_info = ""
				prog_c_info = ""
				prog_e_info = ""
				prog_n_info = ""

				If HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked Then race_a_info = "YES"
				If HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked Then race_b_info = "YES"
				If HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked Then race_n_info = "YES"
				If HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked Then race_p_info = "YES"
				If HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked Then race_w_info = "YES"
				If HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked Then prog_s_info = "YES"
				If HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked Then prog_c_info = "YES"
				If HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked Then prog_e_info = "YES"
				If HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked Then prog_n_info = "YES"

				objTextStream.WriteLine "ARR - HH_MEMB_ARRAY - " & HH_MEMB_ARRAY(ref_number, known_membs)&"~"&HH_MEMB_ARRAY(access_denied, known_membs)&"~"&HH_MEMB_ARRAY(full_name_const, known_membs)&"~"&HH_MEMB_ARRAY(last_name_const, known_membs)&"~"&_
				HH_MEMB_ARRAY(first_name_const, known_membs)&"~"&HH_MEMB_ARRAY(mid_initial, known_membs)&"~"&HH_MEMB_ARRAY(other_names, known_membs)&"~"&HH_MEMB_ARRAY(age, known_membs)&"~"&HH_MEMB_ARRAY(date_of_birth, known_membs)&"~"&HH_MEMB_ARRAY(ssn, known_membs)&"~"&HH_MEMB_ARRAY(ssn_verif, known_membs)&"~"&_
				HH_MEMB_ARRAY(birthdate_verif, known_membs)&"~"&HH_MEMB_ARRAY(gender, known_membs)&"~"&HH_MEMB_ARRAY(race, known_membs)&"~"&HH_MEMB_ARRAY(spoken_lang, known_membs)&"~"&HH_MEMB_ARRAY(written_lang, known_membs)&"~"&HH_MEMB_ARRAY(interpreter, known_membs)&"~"&_
				HH_MEMB_ARRAY(alias_yn, known_membs)&"~"&HH_MEMB_ARRAY(ethnicity_yn, known_membs)&"~"&HH_MEMB_ARRAY(id_verif, known_membs)&"~"&HH_MEMB_ARRAY(rel_to_applcnt, known_membs)&"~"&HH_MEMB_ARRAY(cash_minor, known_membs)&"~"&HH_MEMB_ARRAY(snap_minor, known_membs)&"~"&_
				HH_MEMB_ARRAY(marital_status, known_membs)&"~"&HH_MEMB_ARRAY(spouse_ref, known_membs)&"~"&HH_MEMB_ARRAY(spouse_name, known_membs)&"~"&HH_MEMB_ARRAY(last_grade_completed, known_membs)&"~"&HH_MEMB_ARRAY(citizen, known_membs)&"~"&_
				HH_MEMB_ARRAY(other_st_FS_end_date, known_membs)&"~"&HH_MEMB_ARRAY(in_mn_12_mo, known_membs)&"~"&HH_MEMB_ARRAY(residence_verif, known_membs)&"~"&HH_MEMB_ARRAY(mn_entry_date, known_membs)&"~"&HH_MEMB_ARRAY(former_state, known_membs)&"~"&_
				HH_MEMB_ARRAY(fs_pwe, known_membs)&"~"&HH_MEMB_ARRAY(button_one, known_membs)&"~"&HH_MEMB_ARRAY(button_two, known_membs)&"~"&HH_MEMB_ARRAY(clt_has_sponsor, known_membs)&"~"&HH_MEMB_ARRAY(client_verification, known_membs)&"~"&_
				HH_MEMB_ARRAY(client_verification_details, known_membs)&"~"&HH_MEMB_ARRAY(client_notes, known_membs)&"~"&HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)&"~"&race_a_info&"~"&race_b_info&"~"&race_n_info&"~"&race_p_info&"~"&race_w_info&"~"&prog_s_info&"~"&prog_c_info&"~"&_
				prog_e_info&"~"&prog_n_info&"~"&HH_MEMB_ARRAY(ssn_no_space, known_membs)&"~"&HH_MEMB_ARRAY(edrs_msg, known_membs)&"~"&HH_MEMB_ARRAY(edrs_match, known_membs)&"~"&_
				HH_MEMB_ARRAY(edrs_notes, known_membs)&"~"&HH_MEMB_ARRAY(ignore_person, known_membs)&"~"&HH_MEMB_ARRAY(pers_in_maxis, known_membs)&"~"&HH_MEMB_ARRAY(memb_is_caregiver, known_membs)&"~"&_
                HH_MEMB_ARRAY(cash_request_const, known_membs)&"~"&HH_MEMB_ARRAY(hours_per_week_const, known_membs)&"~"&HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(comply_with_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_needed_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(orientation_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_exempt_const, known_membs)&"~"&HH_MEMB_ARRAY(exemption_reason_const, known_membs)&"~"&HH_MEMB_ARRAY(emps_exemption_code_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(choice_form_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_notes, known_membs)&"~"&HH_MEMB_ARRAY(last_const, known_membs)
			Next

			'Close the object so it can be opened again shortly
			objTextStream.Close

			script_run_lowdown = "* * * INTERVIEW SUMMARY * * *"
			script_run_lowdown = script_run_lowdown & vbCr & "TIME SPENT - "	& timer - start_time & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "MFIP - ORNT - " & MFIP_orientation_assessed_and_completed & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "MFIP - DWP - " & family_cash_program
            script_run_lowdown = script_run_lowdown & vbCr & "FMCA - 01 - " & famliy_cash_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "PROG - CASH - " & cash_other_req_detail
			script_run_lowdown = script_run_lowdown & vbCr & "PROG - SNAP - " & snap_other_req_detail
			script_run_lowdown = script_run_lowdown & vbCr & "PROG - EMER - " & emer_other_req_detail & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "PRE - ATC - " & all_the_clients
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - WHO - " & who_are_we_completing_the_interview_with
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - HOW - " & how_are_we_completing_the_interview
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - ITP - " & interpreter_information
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - LNG - " & interpreter_language
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - AID - " & arep_interview_id_information
			script_run_lowdown = script_run_lowdown & vbCr & "PRE - DET - " & non_applicant_interview_info & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 1 - " & exp_q_1_income_this_month
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 2 - " & exp_q_2_assets_this_month
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - RENT - " & exp_q_3_rent_this_month
			If caf_exp_pay_heat_checkbox = checked 			Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - HEAT - CHECKED"
			If caf_exp_pay_ac_checkbox = checked 			Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - ACON - CHECKED"
			If caf_exp_pay_electricity_checkbox = checked 	Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - ELEC - CHECKED"
			If caf_exp_pay_phone_checkbox = checked 		Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - PHON - CHECKED"
			If caf_exp_pay_none_checkbox = checked 			Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - NONE - CHECKED"
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 3 - UTIL - " & exp_q_4_utilities_this_month
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 4 - " & exp_migrant_seasonal_formworker_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 5 - PREV - " & exp_received_previous_assistance_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 5 - WHEN - " & exp_previous_assistance_when
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 5 - WHER - " & exp_previous_assistance_where
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 5 - WHAT - " & exp_previous_assistance_what
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 6 - PREG - " & exp_pregnant_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - 6 - WHO? - " & exp_pregnant_who & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - INCM - " & intv_app_month_income
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - ASST - " & intv_app_month_asset
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - RENT - " & intv_app_month_housing_expense
			If intv_exp_pay_heat_checkbox = checked 		Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - HEAT - CHECKED"
			If intv_exp_pay_ac_checkbox = checked 			Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - ACON - CHECKED"
			If intv_exp_pay_electricity_checkbox = checked 	Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - ELEC - CHECKED"
			If intv_exp_pay_phone_checkbox = checked 		Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - PHON - CHECKED"
			If intv_exp_pay_none_checkbox = checked 		Then script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - NONE - CHECKED"
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - ID - " & id_verif_on_file
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - 89 - " & snap_active_in_other_state
			script_run_lowdown = script_run_lowdown & vbCr & "EXP - INTVW - EXP - " & last_snap_was_exp & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - STR - " & resi_addr_street_full
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - CIT - " & resi_addr_city
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - STA - " & resi_addr_state
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - ZIP - " & resi_addr_zip

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - RES - " & reservation_yn
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - NAM - " & reservation_name

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - HML - " & homeless_yn

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - LIV - " & living_situation & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - MAIL - STR - " & mail_addr_street_full
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - MAIL - CIT - " & mail_addr_city
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - MAIL - STA - " & mail_addr_state
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - MAIL - ZIP - " & mail_addr_zip & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - NON - " & phone_one_number
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - TON - " & phone_one_type
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - NTW - " & phone_two_number
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - TTW - " & phone_two_type
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - NTH - " & phone_three_number
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - PHON - TTH - " & phone_three_type & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - DATE - " & address_change_date
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - CNTY - " & resi_addr_county & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - HH - " & household_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - EI - " & earned_income_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - UNEA - " & unearned_income_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - SHEL - " & housing_expanese_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - EXP - " & other_expenses_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - AST - " & assets_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - OTHR - " & other_notes_info
			script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - CHG - " & form_changes_info

			If rights_responsibilities_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - R&R - CHECKED"
			If privacy_practices_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - PRIV - CHECKED"
			If ebt_card_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - EBT - CHECKED"
			If complaints_civil_rights_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - CIV RIGHTS - CHECKED"
			If reporting_resp_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - REPRT RESP - CHECKED"
			If program_info_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - PROG INFO - CHECKED"
			If child_support_info_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - CS INFO - CHECKED"
			If mfip_minor_caregiver_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - MFIP CRGVR - CHECKED"
			If renewal_info_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - RENWL - CHECKED"
			If ievs_info_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - IEVS - CHECKED"
			If appeal_rights_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CASE DETAILS - APPEAL - CHECKED"

			script_run_lowdown = script_run_lowdown & vbCr & "QQ1A - " & qual_question_one
			script_run_lowdown = script_run_lowdown & vbCr & "QQ1M - " & qual_memb_one
			script_run_lowdown = script_run_lowdown & vbCr & "QQ2A - " & qual_question_two
			script_run_lowdown = script_run_lowdown & vbCr & "QQ2M - " & qual_memb_two
			script_run_lowdown = script_run_lowdown & vbCr & "QQ3A - " & qual_question_three
			script_run_lowdown = script_run_lowdown & vbCr & "QQ3M - " & qual_memb_three
			script_run_lowdown = script_run_lowdown & vbCr & "QQ4A - " & qual_question_four
			script_run_lowdown = script_run_lowdown & vbCr & "QQ4M - " & qual_memb_four
			script_run_lowdown = script_run_lowdown & vbCr & "QQ5A - " & qual_question_five
			script_run_lowdown = script_run_lowdown & vbCr & "QQ5M - " & qual_memb_five & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 001 - " & arep_in_MAXIS
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 002 - " & MAXIS_arep_updated
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 003 - " & arep_authorization
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 004 - " & arep_authorized

			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 01 - " & arep_name
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 02 - " & arep_relationship
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 03 - " & arep_phone_number
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 04 - " & arep_addr_street
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 05 - " & arep_addr_city
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 06 - " & arep_addr_state
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 07 - " & arep_addr_zip
			If arep_complete_forms_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "AREP - 08 - CHECKED"
			If arep_get_notices_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "AREP - 09 - CHECKED"
			If arep_use_SNAP_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "AREP - 10 - CHECKED"
			If arep_on_CAF_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "AREP - 11 - CHECKED"
			script_run_lowdown = script_run_lowdown & vbCr & "AREP - 12 - " & arep_action & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 01 - " & MAXIS_arep_name
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 02 - " & MAXIS_arep_relationship
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 03 - " & MAXIS_arep_phone_number
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 04 - " & MAXIS_arep_addr_street
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 05 - " & MAXIS_arep_addr_city
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 06 - " & MAXIS_arep_addr_state
			script_run_lowdown = script_run_lowdown & vbCr & "MX-AREP - 07 - " & MAXIS_arep_addr_zip & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 01 - " & CAF_arep_name
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 02 - " & CAF_arep_relationship
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 03 - " & CAF_arep_phone_number
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 04 - " & CAF_arep_addr_street
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 05 - " & CAF_arep_addr_city
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 06 - " & CAF_arep_addr_state
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 07 - " & CAF_arep_addr_zip
			If CAF_arep_complete_forms_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 08"
			If CAF_arep_get_notices_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 09"
			If CAF_arep_use_SNAP_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 10"
			script_run_lowdown = script_run_lowdown & vbCr & "CAF-AREP - 11 - " & CAF_arep_action & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 01 - " & signature_detail
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 02 - " & signature_person
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 03 - " & signature_date
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 04 - " & second_signature_detail
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 05 - " & second_signature_person
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 06 - " & second_signature_date
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 07 - " & client_signed_verbally_yn
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 08 - " & interview_date & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "ASSESS - 01 - " & exp_snap_approval_date
			script_run_lowdown = script_run_lowdown & vbCr & "ASSESS - 02 - " & exp_snap_delays
			script_run_lowdown = script_run_lowdown & vbCr & "ASSESS - 03 - " & snap_denial_date
			script_run_lowdown = script_run_lowdown & vbCr & "ASSESS - 04 - " & snap_denial_explain
			script_run_lowdown = script_run_lowdown & vbCr & "ASSESS - 05 - " & pend_snap_on_case & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - TOTAL - " & discrepancies_exist
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHONE - 01 - " & disc_no_phone_number
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHONE - 02 - " & disc_phone_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHEXP - 01 - " & disc_yes_phone_no_expense
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHEXP - 02 - " & disc_yes_phone_no_expense_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHEXP - 03 - " & disc_no_phone_yes_expense
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - PHEXP - 04 - " & disc_no_phone_yes_expense_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - HOMLS - 01 - " & disc_homeless_no_mail_addr
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - HOMLS - 02 - " & disc_homeless_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - OTOCO - 01 - " & disc_out_of_county
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - OTOCO - 02 - " & disc_out_of_county_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - HOUS$ - 01 - " & disc_rent_amounts
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - HOUS$ - 02 - " & disc_rent_amounts_confirmation
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - UTIL$ - 01 - " & disc_utility_amounts
			script_run_lowdown = script_run_lowdown & vbCr & "CLAR - UTIL$ - 02 - " & disc_utility_amounts_confirmation & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 01 - " & expedited_determination_completed
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 02 - " & expedited_screening
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 03 - " & calculated_low_income_asset_test
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 04 - " & calculated_resources_less_than_expenses_test
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 05 - " & is_elig_XFS
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 06 - " & case_assesment_text
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 07 - " & next_steps_one
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 08 - " & next_steps_two
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 09 - " & next_steps_three
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 10 - " & next_steps_four
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 11 - " & caf_1_resources
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 12 - " & caf_1_expenses
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 13 - " & applicant_id_on_file_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 14 - " & applicant_id_through_SOLQ
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 15 - " & approval_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 16 - " & day_30_from_application
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 17 - " & delay_explanation
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 18 - " & postponed_verifs_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 19 - " & list_postponed_verifs
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 20 - " & first_time_in_exp_det
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 21 - " & income_review_completed
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 22 - " & assets_review_completed
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 23 - " & shel_review_completed
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - 24 - " & note_calculation_detail & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - INCM - 01 - " & determined_income
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - INCM - 02 - " & jobs_income_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - INCM - 03 - " & busi_income_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - INCM - 04 - " & unea_income_yn
			For each_item = 0 to UBound(EXP_JOBS_ARRAY, 2)
				script_run_lowdown = script_run_lowdown & vbCr & "ARR - EXP_JOBS_ARRAY - " & EXP_JOBS_ARRAY(jobs_employee_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_employer_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_wage_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_hours_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_frequency_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_monthly_pay_const, each_item)&"~"&EXP_JOBS_ARRAY(jobs_notes_const, each_item)
			Next
			For each_item = 0 to UBound(EXP_BUSI_ARRAY, 2)
				script_run_lowdown = script_run_lowdown & vbCr & "ARR - EXP_BUSI_ARRAY - " & EXP_BUSI_ARRAY(busi_owner_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_info_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_monthly_earnings_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_annual_earnings_const, each_item)&"~"&EXP_BUSI_ARRAY(busi_notes_const, each_item)
			Next
			For each_item = 0 to UBound(EXP_UNEA_ARRAY, 2)
				script_run_lowdown = script_run_lowdown & vbCr & "ARR - EXP_UNEA_ARRAY - " & EXP_UNEA_ARRAY(unea_owner_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_info_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_monthly_earnings_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_weekly_earnings_const, each_item)&"~"&EXP_UNEA_ARRAY(unea_notes_const, each_item)
			Next
			script_run_lowdown = script_run_lowdown & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - ASST - 01 - " & determined_assets
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - ASST - 02 - " & cash_amount_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - ASST - 03 - " & bank_account_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - ASST - 04 - " & cash_amount
			For each_item = 0 to UBound(EXP_ACCT_ARRAY, 2)
				script_run_lowdown = script_run_lowdown & vbCr & "ARR - EXP_ACCT_ARRAY - " & EXP_ACCT_ARRAY(account_type_const, each_item)&"~"&EXP_ACCT_ARRAY(account_owner_const, each_item)&"~"&EXP_ACCT_ARRAY(bank_name_const, each_item)&"~"&EXP_ACCT_ARRAY(account_amount_const, each_item)&"~"&EXP_ACCT_ARRAY(account_notes_const, each_item)
			Next
			script_run_lowdown = script_run_lowdown & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 01 - " & determined_shel
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 02 - " & rent_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 03 - " & lot_rent_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 04 - " & mortgage_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 05 - " & insurance_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 06 - " & tax_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 07 - " & room_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - SHEL - 08 - " & garage_amount & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 01 - " & determined_utilities
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 02 - " & heat_expense
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 03 - " & ac_expense
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 04 - " & electric_expense
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 05 - " & phone_expense
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 06 - " & none_expense
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - HEST - 07 - " & all_utilities
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - RESOURCES - " & calculated_resources
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - EXPENSES - " & calculated_expenses & vbCr & vbCr


			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 01 - " & other_snap_state
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 02 - " & other_state_reported_benefit_end_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 03 - " & other_state_benefits_openended
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 04 - " & other_state_contact_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 05 - " & other_state_verified_benefit_end_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 06 - " & mn_elig_begin_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - OUTSTATE - 07 - " & action_due_to_out_of_state_benefits & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 01 - " & case_has_previously_postponed_verifs_that_prevent_exp_snap
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 02 - " & prev_post_verif_assessment_done
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 03 - " & previous_CAF_datestamp
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 04 - " & previous_expedited_package
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 05 - " & prev_verifs_mandatory_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 06 - " & prev_verif_list
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 07 - " & curr_verifs_postponed_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 08 - " & ongoing_snap_approved_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - PSTPND - 09 - " & prev_post_verifs_recvd_yn & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 01 - " & delay_action_due_to_faci
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 02 - " & deny_snap_due_to_faci
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 03 - " & faci_review_completed
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 04 - " & facility_name
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 05 - " & snap_inelig_faci_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 06 - " & faci_entry_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 07 - " & faci_release_date
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 08 - " & release_date_unknown_checkbox
			script_run_lowdown = script_run_lowdown & vbCr & "EXPDET - FACI - 09 - " & release_within_30_days_yn & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "VERIFS - " & verifs_selected
			script_run_lowdown = script_run_lowdown & vbCr & "VRFDTE - " & verif_req_form_sent_date

			If number_verifs_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "NUMBER VERIFS - CHECKED"
			If verifs_postponed_checkbox = checked Then script_run_lowdown = script_run_lowdown & vbCr & "POSTPONE VERIFS - CHECKED"
            If verif_snap_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_snap_checkbox - CHECKED"
            If verif_cash_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_cash_checkbox - CHECKED"
            If verif_mfip_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_mfip_checkbox - CHECKED"
            If verif_dwp_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_dwp_checkbox - CHECKED"
            If verif_msa_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_msa_checkbox - CHECKED"
            If verif_ga_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_ga_checkbox - CHECKED"
            If verif_grh_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_grh_checkbox - CHECKED"
            If verif_emer_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_emer_checkbox - CHECKED"
            If verif_hc_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "verif_hc_checkbox - CHECKED" & vbCr & vbCr

			For known_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
				' objTextStream.WriteLine "ARR - ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_last_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_first_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_other_names, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_dob, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_gender, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_former_state, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_citizen, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_written_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_notes, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				race_a_info = ""
				race_b_info = ""
				race_n_info = ""
				race_p_info = ""
				race_w_info = ""
				prog_s_info = ""
				prog_c_info = ""
				prog_e_info = ""
				prog_n_info = ""

				If HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked Then race_a_info = "YES"
				If HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked Then race_b_info = "YES"
				If HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked Then race_n_info = "YES"
				If HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked Then race_p_info = "YES"
				If HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked Then race_w_info = "YES"
				If HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked Then prog_s_info = "YES"
				If HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked Then prog_c_info = "YES"
				If HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked Then prog_e_info = "YES"
				If HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked Then prog_n_info = "YES"

				script_run_lowdown = script_run_lowdown & vbCr & "ARR - HH_MEMB_ARRAY - " & HH_MEMB_ARRAY(ref_number, known_membs)&"~"&HH_MEMB_ARRAY(access_denied, known_membs)&"~"&HH_MEMB_ARRAY(full_name_const, known_membs)&"~"&HH_MEMB_ARRAY(last_name_const, known_membs)&"~"&_
				HH_MEMB_ARRAY(first_name_const, known_membs)&"~"&HH_MEMB_ARRAY(mid_initial, known_membs)&"~"&HH_MEMB_ARRAY(other_names, known_membs)&"~"&HH_MEMB_ARRAY(age, known_membs)&"~"&HH_MEMB_ARRAY(date_of_birth, known_membs)&"~"&HH_MEMB_ARRAY(ssn, known_membs)&"~"&HH_MEMB_ARRAY(ssn_verif, known_membs)&"~"&_
				HH_MEMB_ARRAY(birthdate_verif, known_membs)&"~"&HH_MEMB_ARRAY(gender, known_membs)&"~"&HH_MEMB_ARRAY(race, known_membs)&"~"&HH_MEMB_ARRAY(spoken_lang, known_membs)&"~"&HH_MEMB_ARRAY(written_lang, known_membs)&"~"&HH_MEMB_ARRAY(interpreter, known_membs)&"~"&_
				HH_MEMB_ARRAY(alias_yn, known_membs)&"~"&HH_MEMB_ARRAY(ethnicity_yn, known_membs)&"~"&HH_MEMB_ARRAY(id_verif, known_membs)&"~"&HH_MEMB_ARRAY(rel_to_applcnt, known_membs)&"~"&HH_MEMB_ARRAY(cash_minor, known_membs)&"~"&HH_MEMB_ARRAY(snap_minor, known_membs)&"~"&_
				HH_MEMB_ARRAY(marital_status, known_membs)&"~"&HH_MEMB_ARRAY(spouse_ref, known_membs)&"~"&HH_MEMB_ARRAY(spouse_name, known_membs)&"~"&HH_MEMB_ARRAY(last_grade_completed, known_membs)&"~"&HH_MEMB_ARRAY(citizen, known_membs)&"~"&_
				HH_MEMB_ARRAY(other_st_FS_end_date, known_membs)&"~"&HH_MEMB_ARRAY(in_mn_12_mo, known_membs)&"~"&HH_MEMB_ARRAY(residence_verif, known_membs)&"~"&HH_MEMB_ARRAY(mn_entry_date, known_membs)&"~"&HH_MEMB_ARRAY(former_state, known_membs)&"~"&_
				HH_MEMB_ARRAY(fs_pwe, known_membs)&"~"&HH_MEMB_ARRAY(button_one, known_membs)&"~"&HH_MEMB_ARRAY(button_two, known_membs)&"~"&HH_MEMB_ARRAY(clt_has_sponsor, known_membs)&"~"&HH_MEMB_ARRAY(client_verification, known_membs)&"~"&_
				HH_MEMB_ARRAY(client_verification_details, known_membs)&"~"&HH_MEMB_ARRAY(client_notes, known_membs)&"~"&HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)&"~"&race_a_info&"~"&race_b_info&"~"&race_n_info&"~"&race_p_info&"~"&race_w_info&"~"&prog_s_info&"~"&prog_c_info&"~"&_
				prog_e_info&"~"&prog_n_info&"~"&HH_MEMB_ARRAY(ssn_no_space, known_membs)&"~"&HH_MEMB_ARRAY(edrs_msg, known_membs)&"~"&HH_MEMB_ARRAY(edrs_match, known_membs)&"~"&_
                HH_MEMB_ARRAY(edrs_notes, known_membs)&"~"&HH_MEMB_ARRAY(ignore_person, known_membs)&"~"&HH_MEMB_ARRAY(pers_in_maxis, known_membs)&"~"&HH_MEMB_ARRAY(memb_is_caregiver, known_membs)&"~"&_
                HH_MEMB_ARRAY(cash_request_const, known_membs)&"~"&HH_MEMB_ARRAY(hours_per_week_const, known_membs)&"~"&HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(comply_with_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_needed_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(orientation_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_exempt_const, known_membs)&"~"&HH_MEMB_ARRAY(exemption_reason_const, known_membs)&"~"&HH_MEMB_ARRAY(emps_exemption_code_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(choice_form_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_notes, known_membs)&"~"&HH_MEMB_ARRAY(last_const, known_membs) & vbCr & vbCr
			Next

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run
function restore_your_work(vars_filled)

	'Now determines name of file
	save_your_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"

	With (CreateObject("Scripting.FileSystemObject"))

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(save_your_work_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

			If pull_variables = vbYes Then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(save_your_work_path, ForReading)

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				saved_caf_details = split(every_line_in_text_file, vbNewLine)
				vars_filled = TRUE

				array_counters = 0
				known_membs = 0
				known_jobs = 0
				known_exp_jobs = 0
				known_exp_busi = 0
				known_exp_unea = 0
				known_exp_acct = 0
				For Each text_line in saved_caf_details
					' MsgBox "~" & left(text_line, 9) & "~" & vbCr & text_line
					' MsgBox text_line
					If left(text_line, 4) = "TIME" Then add_to_time = right(text_line, len(text_line) - 13)
					add_to_time = trim(add_to_time)
					If IsNumeric(add_to_time) = True Then add_to_time = add_to_time * 1

					If left(text_line, 10) = "CAF - DATE" Then CAF_datestamp = Mid(text_line, 14)
                    If left(text_line, 11) = "MFIP - ORNT" Then MFIP_orientation_assessed_and_completed = Mid(text_line, 15)
                    If UCase(MFIP_orientation_assessed_and_completed) = "TRUE" Then MFIP_orientation_assessed_and_completed = True
                    If UCase(MFIP_orientation_assessed_and_completed) = "FALSE" Then MFIP_orientation_assessed_and_completed = False
                    If left(text_line, 10) = "MFIP - DWP" Then family_cash_program = Mid(text_line, 14)
                    If left(text_line, 9) = "FMCA - 01" Then famliy_cash_notes = Mid(text_line, 13)

					If left(text_line, 11) = "PROG - CASH" Then cash_other_req_detail = Mid(text_line, 15)
					If left(text_line, 11) = "PROG - SNAP" Then snap_other_req_detail = Mid(text_line, 15)
					If left(text_line, 11) = "PROG - EMER" Then emer_other_req_detail = Mid(text_line, 15)
					If left(text_line, 17) = "CASH PROG CHECKED" Then CASH_on_CAF_checkbox = checked
					If left(text_line, 17) = "SNAP PROG CHECKED" Then SNAP_on_CAF_checkbox = checked
					If left(text_line, 17) = "EMER PROG CHECKED" Then EMER_on_CAF_checkbox = checked

					If left(text_line, 11) = "CASH - TYPE" Then type_of_cash = Mid(text_line, 15)
					If left(text_line, 11) = "PROC - CASH" Then the_process_for_cash = Mid(text_line, 15)
					If left(text_line, 11) = "CASH - RVMO" Then next_cash_revw_mo = Mid(text_line, 15)
					If left(text_line, 11) = "CASH - RVYR" Then next_cash_revw_yr = Mid(text_line, 15)

					If left(text_line, 11) = "PROC - SNAP" Then the_process_for_snap = Mid(text_line, 15)
					If left(text_line, 11) = "SNAP - RVMO" Then next_snap_revw_mo = Mid(text_line, 15)
					If left(text_line, 11) = "SNAP - RVYR" Then next_snap_revw_yr = Mid(text_line, 15)

					If left(text_line, 11) = "EMER - TYPE" Then type_of_emer = Mid(text_line, 15)
					If left(text_line, 11) = "PROC - EMER" Then the_process_for_emer = Mid(text_line, 15)

					If left(text_line, 9) = "PRE - WHO" Then who_are_we_completing_the_interview_with = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - HOW" Then how_are_we_completing_the_interview = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - ATC" Then all_the_clients = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - ITP" Then interpreter_information = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - LNG" Then interpreter_language = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - AID" Then arep_interview_id_information = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - DET" Then non_applicant_interview_info = Mid(text_line, 13)

					If left(text_line, 7) = "EXP - 1" Then exp_q_1_income_this_month = Mid(text_line, 11)
					If left(text_line, 7) = "EXP - 2" Then exp_q_2_assets_this_month = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 3 - RENT" Then exp_q_3_rent_this_month = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - HEAT" Then caf_exp_pay_heat_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - ACON" Then caf_exp_pay_ac_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - ELEC" Then caf_exp_pay_electricity_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - PHON" Then caf_exp_pay_phone_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - NONE" Then caf_exp_pay_none_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - UTIL" Then exp_q_4_utilities_this_month = Mid(text_line, 18)
					If left(text_line, 7) = "EXP - 4" Then exp_migrant_seasonal_formworker_yn = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 5 - PREV" Then exp_received_previous_assistance_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHEN" Then exp_previous_assistance_when = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHER" Then exp_previous_assistance_where = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHAT" Then exp_previous_assistance_what = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - PREG" Then exp_pregnant_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - WHO?" Then exp_pregnant_who = Mid(text_line, 18)

					If left(text_line, 18) = "EXP - INTVW - INCM" Then intv_app_month_income = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - ASST" Then intv_app_month_asset = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - RENT" Then intv_app_month_housing_expense = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - HEAT" Then intv_exp_pay_heat_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - ACON" Then intv_exp_pay_ac_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - ELEC" Then intv_exp_pay_electricity_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - PHON" Then intv_exp_pay_phone_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - NONE" Then intv_exp_pay_none_checkbox = checked
					If left(text_line, 16) = "EXP - INTVW - ID" Then id_verif_on_file = Mid(text_line, 20)
					If left(text_line, 16) = "EXP - INTVW - 89" Then snap_active_in_other_state = Mid(text_line, 20)
					If left(text_line, 17) = "EXP - INTVW - EXP" Then last_snap_was_exp = Mid(text_line, 21)

					If left(text_line, 3) = "ADR" Then
						' MsgBox "~" & mid(text_line, 7, 10) & "~"
						If mid(text_line, 7, 10) = "RESI - STR" Then resi_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - CIT" Then resi_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - STA" Then resi_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - ZIP" Then resi_addr_zip = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - RES" Then reservation_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - NAM" Then reservation_name = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - HML" Then homeless_yn = MID(text_line, 20)
						If mid(text_line, 7, 10) = "RESI - LIV" Then living_situation = MID(text_line, 20)

						If mid(text_line, 7, 10) = "MAIL - STR" Then mail_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - CIT" Then mail_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - STA" Then mail_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - ZIP" Then mail_addr_zip = MID(text_line, 20)

						If mid(text_line, 7, 10) = "PHON - NON" Then phone_one_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TON" Then phone_one_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTW" Then phone_two_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTW" Then phone_two_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTH" Then phone_three_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTH" Then phone_three_type = MID(text_line, 20)

						If mid(text_line, 7, 4) = "DATE" Then address_change_date = MID(text_line, 14)
						If mid(text_line, 7, 4) = "CNTY" Then resi_addr_county = MID(text_line, 14)

					End If

					If left(text_line, 17) = "CASE DETAILS - HH" Then household_info = Mid(text_line, 21)
					If left(text_line, 17) = "CASE DETAILS - EI" Then earned_income_info = Mid(text_line, 21)
					If left(text_line, 19) = "CASE DETAILS - UNEA" Then unearned_income_info = Mid(text_line, 23)
					If left(text_line, 19) = "CASE DETAILS - SHEL" Then housing_expanese_info = Mid(text_line, 23)
					If left(text_line, 18) = "CASE DETAILS - EXP" Then other_expenses_info = Mid(text_line, 22)
					If left(text_line, 18) = "CASE DETAILS - AST" Then assets_info = Mid(text_line, 22)
					If left(text_line, 19) = "CASE DETAILS - OTHR" Then other_notes_info = Mid(text_line, 23)
					If left(text_line, 18) = "CASE DETAILS - CHG" Then form_changes_info = Mid(text_line, 22)

					If text_line = "rights_responsibilities_checkbox" Then rights_responsibilities_checkbox = checked
					If text_line = "privacy_practices_checkbox" Then privacy_practices_checkbox = checked
					If text_line = "ebt_card_checkbox" Then ebt_card_checkbox = checked
					If text_line = "complaints_civil_rights_checkbox" Then complaints_civil_rights_checkbox = checked
					If text_line = "reporting_resp_checkbox" Then reporting_resp_checkbox = checked
					If text_line = "program_info_checkbox" Then program_info_checkbox = checked
					If text_line = "child_support_info_checkbox" Then child_support_info_checkbox = checked
					If text_line = "mfip_minor_caregiver_checkbox" Then mfip_minor_caregiver_checkbox = checked
					If text_line = "renewal_info_checkbox" Then renewal_info_checkbox = checked
					If text_line = "ievs_info_checkbox" Then ievs_info_checkbox = checked
					If text_line = "appeal_rights_checkbox" Then appeal_rights_checkbox = checked

					If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ1M" Then qual_memb_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2A" Then qual_question_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2M" Then qual_memb_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3A" Then qual_question_three = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3M" Then qual_memb_three = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4A" Then qual_question_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ4M" Then qual_memb_four = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5A" Then qual_question_five = Mid(text_line, 8)
					If left(text_line, 4) = "QQ5M" Then qual_memb_five = Mid(text_line, 8)

					If left(text_line, 10) = "AREP - 001" Then arep_in_MAXIS = Mid(text_line, 14)
					If left(text_line, 10) = "AREP - 002" Then MAXIS_arep_updated = Mid(text_line, 14)
					If left(text_line, 10) = "AREP - 003" Then arep_authorization = Mid(text_line, 14)
					If left(text_line, 10) = "AREP - 004" Then arep_authorized = Mid(text_line, 14)

					If left(text_line, 9) = "AREP - 01" Then arep_name = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 02" Then arep_relationship = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 03" Then arep_phone_number = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 04" Then arep_addr_street = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 05" Then arep_addr_city = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 06" Then arep_addr_state = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 07" Then arep_addr_zip = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 08" Then arep_complete_forms_checkbox = checked
					If left(text_line, 9) = "AREP - 09" Then arep_get_notices_checkbox = checked
					If left(text_line, 9) = "AREP - 10" Then arep_use_SNAP_checkbox = checked
					If left(text_line, 9) = "AREP - 11" Then arep_on_CAF_checkbox = checked
					If left(text_line, 9) = "AREP - 12" Then arep_action = Mid(text_line, 13)

					If left(text_line, 12) = "MX-AREP - 01" Then MAXIS_arep_name = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 02" Then MAXIS_arep_relationship = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 03" Then MAXIS_arep_phone_number = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 04" Then MAXIS_arep_addr_street = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 05" Then MAXIS_arep_addr_city = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 06" Then MAXIS_arep_addr_state = Mid(text_line, 16)
					If left(text_line, 12) = "MX-AREP - 07" Then MAXIS_arep_addr_zip = Mid(text_line, 16)

					If left(text_line, 13) = "CAF-AREP - 01" Then CAF_arep_name = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 02" Then CAF_arep_relationship = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 03" Then CAF_arep_phone_number = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 04" Then CAF_arep_addr_street = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 05" Then CAF_arep_addr_city = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 06" Then CAF_arep_addr_state = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 07" Then CAF_arep_addr_zip = Mid(text_line, 17)
					If left(text_line, 13) = "CAF-AREP - 08" Then CAF_arep_complete_forms_checkbox = checked
					If left(text_line, 13) = "CAF-AREP - 09" Then CAF_arep_get_notices_checkbox = checked
					If left(text_line, 13) = "CAF-AREP - 10" Then CAF_arep_use_SNAP_checkbox = checked
					If left(text_line, 13) = "CAF-AREP - 11" Then CAF_arep_action = Mid(text_line, 17)

					If left(text_line, 8) = "SIG - 01" Then signature_detail = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 02" Then signature_person = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 03" Then signature_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 04" Then second_signature_detail = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 05" Then second_signature_person = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 06" Then second_signature_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 07" Then client_signed_verbally_yn = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 08" Then interview_date = Mid(text_line, 12)

					If left(text_line, 11) = "ASSESS - 01" Then exp_snap_approval_date = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 02" Then exp_snap_delays = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 03" Then snap_denial_date = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 04" Then snap_denial_explain = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 05" Then pend_snap_on_case = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 06" Then family_cash_case_yn = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 07" Then absent_parent_yn = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 08" Then relative_caregiver_yn = Mid(text_line, 15)
					If left(text_line, 11) = "ASSESS - 09" Then minor_caregiver_yn = Mid(text_line, 15)

					If left(text_line, 12) = "CLAR - TOTAL" Then read_disc = UCASE(text_line)
					If Instr(read_disc, "TRUE") Then discrepancies_exist = True
					If Instr(read_disc, "FALSE") Then discrepancies_exist = False
					If left(text_line, 17) = "CLAR - PHONE - 01" Then disc_no_phone_number = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - PHONE - 02" Then disc_phone_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - PHEXP - 01" Then disc_yes_phone_no_expense = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - PHEXP - 02" Then disc_yes_phone_no_expense_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - PHEXP - 03" Then disc_no_phone_yes_expense = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - PHEXP - 04" Then disc_no_phone_yes_expense_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - HOMLS - 01" Then disc_homeless_no_mail_addr = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - HOMLS - 02" Then disc_homeless_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - OTOCO - 01" Then disc_out_of_county = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - OTOCO - 02" Then disc_out_of_county_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - HOUS$ - 01" Then disc_rent_amounts = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - HOUS$ - 02" Then disc_rent_amounts_confirmation = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - UTIL$ - 01" Then disc_utility_amounts = Mid(text_line, 21)
					If left(text_line, 17) = "CLAR - UTIL$ - 02" Then disc_utility_amounts_confirmation = Mid(text_line, 21)


					If left(text_line, 11) = "EXPDET - 01" Then expedited_determination_completed = Mid(text_line, 15)
					If UCASE(expedited_determination_completed) = "TRUE" Then expedited_determination_completed = True
					If UCASE(expedited_determination_completed) = "FALSE" Then expedited_determination_completed = False
					If left(text_line, 11) = "EXPDET - 02" Then expedited_screening = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 03" Then calculated_low_income_asset_test = Mid(text_line, 15)
					If UCASE(calculated_low_income_asset_test) = "TRUE" Then calculated_low_income_asset_test = True
					If UCASE(calculated_low_income_asset_test) = "FALSE" Then calculated_low_income_asset_test = False
					If left(text_line, 11) = "EXPDET - 04" Then calculated_resources_less_than_expenses_test = Mid(text_line, 15)
					If UCASE(calculated_resources_less_than_expenses_test) = "TRUE" Then calculated_resources_less_than_expenses_test = True
					If UCASE(calculated_resources_less_than_expenses_test) = "FALSE" Then calculated_resources_less_than_expenses_test = False
					If left(text_line, 11) = "EXPDET - 05" Then is_elig_XFS = Mid(text_line, 15)
					If UCASE(is_elig_XFS) = "TRUE" Then is_elig_XFS = True
					If UCASE(is_elig_XFS) = "FALSE" Then is_elig_XFS = False
					If left(text_line, 11) = "EXPDET - 06" Then case_assesment_text = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 07" Then next_steps_one = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 08" Then next_steps_two = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 09" Then next_steps_three = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 10" Then next_steps_four = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 11" Then caf_1_resources = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 12" Then caf_1_expenses = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 13" Then applicant_id_on_file_yn = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 14" Then applicant_id_through_SOLQ = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 15" Then approval_date = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 16" Then day_30_from_application = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 17" Then delay_explanation = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 18" Then postponed_verifs_yn = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 19" Then list_postponed_verifs = Mid(text_line, 15)
					If left(text_line, 11) = "EXPDET - 20" Then first_time_in_exp_det = Mid(text_line, 15)
					If UCASE(first_time_in_exp_det) = "TRUE" Then first_time_in_exp_det = True
					If UCASE(first_time_in_exp_det) = "FALSE" Then first_time_in_exp_det = False
					If left(text_line, 11) = "EXPDET - 21" Then income_review_completed = Mid(text_line, 15)
					If UCASE(income_review_completed) = "TRUE" Then income_review_completed = True
					If UCASE(income_review_completed) = "FALSE" Then income_review_completed = False
					If left(text_line, 11) = "EXPDET - 22" Then assets_review_completed = Mid(text_line, 15)
					If UCASE(assets_review_completed) = "TRUE" Then assets_review_completed = True
					If UCASE(assets_review_completed) = "FALSE" Then assets_review_completed = False
					If left(text_line, 11) = "EXPDET - 23" Then shel_review_completed = Mid(text_line, 15)
					If UCASE(shel_review_completed) = "TRUE" Then shel_review_completed = True
					If UCASE(shel_review_completed) = "FALSE" Then shel_review_completed = False
					If left(text_line, 11) = "EXPDET - 24" Then note_calculation_detail = Mid(text_line, 15)
					If UCASE(note_calculation_detail) = "TRUE" Then note_calculation_detail = True
					If UCASE(note_calculation_detail) = "FALSE" Then note_calculation_detail = False

					If left(text_line, 18) = "EXPDET - INCM - 01" Then determined_income = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - INCM - 02" Then jobs_income_yn = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - INCM - 03" Then busi_income_yn = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - INCM - 04" Then unea_income_yn = Mid(text_line, 22)


					If left(text_line, 18) = "EXPDET - ASST - 01" Then determined_assets = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - ASST - 02" Then cash_amount_yn = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - ASST - 03" Then bank_account_yn = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - ASST - 04" Then cash_amount = Mid(text_line, 22)


					If left(text_line, 18) = "EXPDET - SHEL - 01" Then determined_shel = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 02" Then rent_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 03" Then lot_rent_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 04" Then mortgage_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 05" Then insurance_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 06" Then tax_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 07" Then room_amount = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - SHEL - 08" Then garage_amount = Mid(text_line, 22)

					If left(text_line, 18) = "EXPDET - HEST - 01" Then determined_utilities = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - HEST - 02" Then heat_expense = Mid(text_line, 22)
					If UCASE(heat_expense) = "TRUE" Then heat_expense = True
					If UCASE(heat_expense) = "FALSE" Then heat_expense = False
					If left(text_line, 18) = "EXPDET - HEST - 03" Then ac_expense = Mid(text_line, 22)
					If UCASE(ac_expense) = "TRUE" Then ac_expense = True
					If UCASE(ac_expense) = "FALSE" Then ac_expense = False
					If left(text_line, 18) = "EXPDET - HEST - 04" Then electric_expense = Mid(text_line, 22)
					If UCASE(electric_expense) = "TRUE" Then electric_expense = True
					If UCASE(electric_expense) = "FALSE" Then electric_expense = False
					If left(text_line, 18) = "EXPDET - HEST - 05" Then phone_expense = Mid(text_line, 22)
					If UCASE(phone_expense) = "TRUE" Then phone_expense = True
					If UCASE(phone_expense) = "FALSE" Then phone_expense = False
					If left(text_line, 18) = "EXPDET - HEST - 06" Then none_expense = Mid(text_line, 22)
					If UCASE(none_expense) = "TRUE" Then none_expense = True
					If UCASE(none_expense) = "FALSE" Then none_expense = False
					If left(text_line, 18) = "EXPDET - HEST - 07" Then all_utilities = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - RESOURCES" Then calculated_resources = Mid(text_line, 22)
					If left(text_line, 17) = "EXPDET - EXPENSES" Then calculated_expenses = Mid(text_line, 21)

					If left(text_line, 22) = "EXPDET - OUTSTATE - 01" Then other_snap_state = Mid(text_line, 26)
					If left(text_line, 22) = "EXPDET - OUTSTATE - 02" Then other_state_reported_benefit_end_date = Mid(text_line, 26)
					If left(text_line, 22) = "EXPDET - OUTSTATE - 03" Then other_state_benefits_openended = Mid(text_line, 26)
					If UCASE(other_state_benefits_openended) = "TRUE" Then other_state_benefits_openended = True
					If UCASE(other_state_benefits_openended) = "FALSE" Then other_state_benefits_openended = False
					If left(text_line, 22) = "EXPDET - OUTSTATE - 04" Then other_state_contact_yn = Mid(text_line, 26)
					If left(text_line, 22) = "EXPDET - OUTSTATE - 05" Then other_state_verified_benefit_end_date = Mid(text_line, 26)
					If left(text_line, 22) = "EXPDET - OUTSTATE - 06" Then mn_elig_begin_date = Mid(text_line, 26)
					If left(text_line, 22) = "EXPDET - OUTSTATE - 07" Then action_due_to_out_of_state_benefits = Mid(text_line, 26)

					If left(text_line, 20) = "EXPDET - PSTPND - 01" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = Mid(text_line, 24)
					If UCASE(case_has_previously_postponed_verifs_that_prevent_exp_snap) = "TRUE" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = True
					If UCASE(case_has_previously_postponed_verifs_that_prevent_exp_snap) = "FALSE" Then case_has_previously_postponed_verifs_that_prevent_exp_snap = False
					If left(text_line, 20) = "EXPDET - PSTPND - 02" Then prev_post_verif_assessment_done = Mid(text_line, 24)
					If UCASE(prev_post_verif_assessment_done) = "TRUE" Then prev_post_verif_assessment_done = True
					If UCASE(prev_post_verif_assessment_done) = "FALSE" Then prev_post_verif_assessment_done = False
					If left(text_line, 20) = "EXPDET - PSTPND - 03" Then previous_CAF_datestamp = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 04" Then previous_expedited_package = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 05" Then prev_verifs_mandatory_yn = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 06" Then prev_verif_list = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 07" Then curr_verifs_postponed_yn = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 08" Then ongoing_snap_approved_yn = Mid(text_line, 24)
					If left(text_line, 20) = "EXPDET - PSTPND - 09" Then prev_post_verifs_recvd_yn = Mid(text_line, 24)

					If left(text_line, 18) = "EXPDET - FACI - 01" Then delay_action_due_to_faci = Mid(text_line, 22)
					If UCASE(delay_action_due_to_faci) = "TRUE" Then delay_action_due_to_faci = True
					If UCASE(delay_action_due_to_faci) = "FALSE" Then delay_action_due_to_faci = False
					If left(text_line, 18) = "EXPDET - FACI - 02" Then deny_snap_due_to_faci = Mid(text_line, 22)
					If UCASE(deny_snap_due_to_faci) = "TRUE" Then deny_snap_due_to_faci = True
					If UCASE(deny_snap_due_to_faci) = "FALSE" Then deny_snap_due_to_faci = False
					If left(text_line, 18) = "EXPDET - FACI - 03" Then faci_review_completed = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - FACI - 04" Then facility_name = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - FACI - 05" Then snap_inelig_faci_yn = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - FACI - 06" Then faci_entry_date = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - FACI - 07" Then faci_release_date = Mid(text_line, 22)
					If left(text_line, 18) = "EXPDET - FACI - 08" Then release_date_unknown_checkbox = checked
					If left(text_line, 18) = "EXPDET - FACI - 09" Then release_within_30_days_yn = Mid(text_line, 22)


					If left(text_line, 6) = "VERIFS" Then verifs_selected = Mid(text_line, 10)
					If left(text_line, 6) = "VRFDTE" Then verif_req_form_sent_date = Mid(text_line, 10)

					If text_line = "NUMBER VERIFS" Then number_verifs_checkbox = checked
					If text_line = "POSTPONE VERIFS" Then verifs_postponed_checkbox = checked
                    If text_line = "verif_snap_checkbox" Then verif_snap_checkbox = checked
                    If text_line = "verif_cash_checkbox" Then verif_cash_checkbox = checked
                    If text_line = "verif_mfip_checkbox" Then verif_mfip_checkbox = checked
                    If text_line = "verif_dwp_checkbox" Then verif_dwp_checkbox = checked
                    If text_line = "verif_msa_checkbox" Then verif_msa_checkbox = checked
                    If text_line = "verif_ga_checkbox" Then verif_ga_checkbox = checked
                    If text_line = "verif_grh_checkbox" Then verif_grh_checkbox = checked
                    If text_line = "verif_emer_checkbox" Then verif_emer_checkbox = checked
                    If text_line = "verif_hc_checkbox" Then verif_hc_checkbox = checked

					If left(text_line, 3) = "ARR" Then
						If MID(text_line, 7, 13) = "HH_MEMB_ARRAY" Then
							array_info = Mid(text_line, 23)
							array_info = split(array_info, "~")
							ReDim Preserve HH_MEMB_ARRAY(last_const, known_membs)
							HH_MEMB_ARRAY(ref_number, known_membs)					= array_info(0)
							HH_MEMB_ARRAY(access_denied, known_membs)				= array_info(1)
							HH_MEMB_ARRAY(full_name_const, known_membs)				= array_info(2)
							HH_MEMB_ARRAY(last_name_const, known_membs)				= array_info(3)
							HH_MEMB_ARRAY(first_name_const, known_membs)			= array_info(4)
							HH_MEMB_ARRAY(mid_initial, known_membs)					= array_info(5)
							HH_MEMB_ARRAY(other_names, known_membs)					= array_info(6)
							HH_MEMB_ARRAY(age, known_membs)							= array_info(7)
							' MsgBox "~" & HH_MEMB_ARRAY(age, known_membs) & "~"
							If HH_MEMB_ARRAY(age, known_membs) = "" Then HH_MEMB_ARRAY(age, known_membs) = 0
							HH_MEMB_ARRAY(age, known_membs) = HH_MEMB_ARRAY(age, known_membs) * 1
							HH_MEMB_ARRAY(date_of_birth, known_membs)				= array_info(8)
							HH_MEMB_ARRAY(ssn, known_membs)							= array_info(9)
							HH_MEMB_ARRAY(ssn_verif, known_membs)					= array_info(10)
							HH_MEMB_ARRAY(birthdate_verif, known_membs)				= array_info(11)
							HH_MEMB_ARRAY(gender, known_membs)						= array_info(12)
							HH_MEMB_ARRAY(race, known_membs)						= array_info(13)
							HH_MEMB_ARRAY(spoken_lang, known_membs)					= array_info(14)
							HH_MEMB_ARRAY(written_lang, known_membs)				= array_info(15)
							HH_MEMB_ARRAY(interpreter, known_membs)					= array_info(16)
							HH_MEMB_ARRAY(alias_yn, known_membs)					= array_info(17)
							HH_MEMB_ARRAY(ethnicity_yn, known_membs)				= array_info(18)
							HH_MEMB_ARRAY(id_verif, known_membs)					= array_info(19)
							HH_MEMB_ARRAY(rel_to_applcnt, known_membs)				= array_info(20)
							HH_MEMB_ARRAY(cash_minor, known_membs)					= array_info(21)
							HH_MEMB_ARRAY(snap_minor, known_membs)					= array_info(22)
							HH_MEMB_ARRAY(marital_status, known_membs)				= array_info(23)
							HH_MEMB_ARRAY(spouse_ref, known_membs)					= array_info(24)
							HH_MEMB_ARRAY(spouse_name, known_membs)					= array_info(25)
							HH_MEMB_ARRAY(last_grade_completed, known_membs) 		= array_info(26)
							HH_MEMB_ARRAY(citizen, known_membs)						= array_info(27)
							HH_MEMB_ARRAY(other_st_FS_end_date, known_membs) 		= array_info(28)
							HH_MEMB_ARRAY(in_mn_12_mo, known_membs)					= array_info(29)
							HH_MEMB_ARRAY(residence_verif, known_membs)				= array_info(30)
							HH_MEMB_ARRAY(mn_entry_date, known_membs)				= array_info(31)
							HH_MEMB_ARRAY(former_state, known_membs)				= array_info(32)
							HH_MEMB_ARRAY(fs_pwe, known_membs)						= array_info(33)
							HH_MEMB_ARRAY(button_one, known_membs)					= array_info(34)
							HH_MEMB_ARRAY(button_two, known_membs)					= array_info(35)
							HH_MEMB_ARRAY(clt_has_sponsor, known_membs)				= array_info(36)
							HH_MEMB_ARRAY(client_verification, known_membs)			= array_info(37)
							HH_MEMB_ARRAY(client_verification_details, known_membs)	= array_info(38)
							HH_MEMB_ARRAY(client_notes, known_membs)				= array_info(39)
							HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)		= array_info(40)
							If array_info(41) = "YES" Then HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked
							If array_info(42) = "YES" Then HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked
							If array_info(43) = "YES" Then HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked
							If array_info(44) = "YES" Then HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked
							If array_info(45) = "YES" Then HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked
							If array_info(46) = "YES" Then HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked
							If array_info(47) = "YES" Then HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked
							If array_info(48) = "YES" Then HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked
							If array_info(49) = "YES" Then HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked
							HH_MEMB_ARRAY(ssn_no_space, known_membs)				= array_info(50)
							HH_MEMB_ARRAY(edrs_msg, known_membs)					= array_info(51)
							HH_MEMB_ARRAY(edrs_match, known_membs)					= array_info(52)
							HH_MEMB_ARRAY(edrs_notes, known_membs) 					= array_info(53)

                            If UBound(array_info) = 69 Then
                                HH_MEMB_ARRAY(ignore_person, known_membs) 			= array_info(54)
                                HH_MEMB_ARRAY(pers_in_maxis, known_membs) 			= array_info(55)
                                HH_MEMB_ARRAY(last_const, known_membs)				= array_info(56)

                                If UCASE(HH_MEMB_ARRAY(ignore_person, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(ignore_person, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(ignore_person, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(ignore_person, known_membs) = False
                                If UCASE(HH_MEMB_ARRAY(pers_in_maxis, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(pers_in_maxis, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(pers_in_maxis, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(pers_in_maxis, known_membs) = False

                                HH_MEMB_ARRAY(memb_is_caregiver, known_membs)      = array_info(57)
                                If UCASE(HH_MEMB_ARRAY(memb_is_caregiver, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(memb_is_caregiver, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(memb_is_caregiver, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(memb_is_caregiver, known_membs) = False

                                HH_MEMB_ARRAY(cash_request_const, known_membs)      = array_info(58)
                                If UCASE(HH_MEMB_ARRAY(cash_request_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(cash_request_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(cash_request_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(cash_request_const, known_membs) = False
                                HH_MEMB_ARRAY(hours_per_week_const, known_membs)    = array_info(59)
                                If IsNumeric(HH_MEMB_ARRAY(hours_per_week_const, known_membs)) = True Then HH_MEMB_ARRAY(hours_per_week_const, known_membs) = HH_MEMB_ARRAY(hours_per_week_const, known_membs) * 1
                                If trim(HH_MEMB_ARRAY(hours_per_week_const, known_membs)) = "" Then HH_MEMB_ARRAY(hours_per_week_const, known_membs) = 0
                                HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)    = array_info(60)
                                If UCASE(HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(exempt_from_ed_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(exempt_from_ed_const, known_membs) = False
                                HH_MEMB_ARRAY(comply_with_ed_const, known_membs)    = array_info(61)
                                If UCASE(HH_MEMB_ARRAY(comply_with_ed_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(comply_with_ed_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(comply_with_ed_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(comply_with_ed_const, known_membs) = False
                                HH_MEMB_ARRAY(orientation_needed_const, known_membs)= array_info(62)
                                If UCASE(HH_MEMB_ARRAY(orientation_needed_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(orientation_needed_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(orientation_needed_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(orientation_needed_const, known_membs) = False

                                HH_MEMB_ARRAY(orientation_done_const, known_membs)  = array_info(63)
                                If UCASE(HH_MEMB_ARRAY(orientation_done_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(orientation_done_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(orientation_done_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(orientation_done_const, known_membs) = False
                                HH_MEMB_ARRAY(orientation_exempt_const, known_membs)= array_info(64)
                                If UCASE(HH_MEMB_ARRAY(orientation_exempt_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(orientation_exempt_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(orientation_exempt_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(orientation_exempt_const, known_membs) = False
                                HH_MEMB_ARRAY(exemption_reason_const, known_membs)  = array_info(65)
                                HH_MEMB_ARRAY(emps_exemption_code_const, known_membs)= array_info(66)

                                HH_MEMB_ARRAY(choice_form_done_const, known_membs)  = array_info(67)
                                If UCASE(HH_MEMB_ARRAY(choice_form_done_const, known_membs)) = "TRUE" Then HH_MEMB_ARRAY(choice_form_done_const, known_membs) = True
                                If UCASE(HH_MEMB_ARRAY(choice_form_done_const, known_membs)) = "FALSE" Then HH_MEMB_ARRAY(choice_form_done_const, known_membs) = False
                                HH_MEMB_ARRAY(orientation_notes, known_membs)       = array_info(68)
                                HH_MEMB_ARRAY(last_const, known_membs)              = array_info(69)


                            Else
                                HH_MEMB_ARRAY(last_const, known_membs)				= array_info(54)

                                HH_MEMB_ARRAY(pers_in_maxis, known_membs) = False
                                If HH_MEMB_ARRAY(ref_number, known_membs) <> "" Then HH_MEMB_ARRAY(pers_in_maxis, known_membs) = True
                                HH_MEMB_ARRAY(ignore_person, known_membs) = False
                            End If

							known_membs = known_membs + 1
						End If

						If MID(text_line, 7, 14) = "EXP_JOBS_ARRAY" Then
							array_info = Mid(text_line, 24)
							array_info = split(array_info, "~")
							ReDim Preserve EXP_JOBS_ARRAY(jobs_notes_const, known_exp_jobs)

							EXP_JOBS_ARRAY(jobs_employee_const, each_item) 		= array_info(0)
							EXP_JOBS_ARRAY(jobs_employer_const, each_item) 		= array_info(1)
							EXP_JOBS_ARRAY(jobs_wage_const, each_item) 			= array_info(2)
							EXP_JOBS_ARRAY(jobs_hours_const, each_item) 		= array_info(3)
							EXP_JOBS_ARRAY(jobs_frequency_const, each_item) 	= array_info(4)
							EXP_JOBS_ARRAY(jobs_monthly_pay_const, each_item) 	= array_info(5)
							EXP_JOBS_ARRAY(jobs_notes_const, each_item) 		= array_info(6)
							known_exp_jobs = known_exp_jobs + 1
						End If

						If MID(text_line, 7, 14) = "EXP_BUSI_ARRAY" Then
							array_info = Mid(text_line, 24)
							array_info = split(array_info, "~")
							ReDim Preserve EXP_BUSI_ARRAY(busi_notes_const, known_exp_busi)

							EXP_BUSI_ARRAY(busi_owner_const, each_item) 			= array_info(0)
							EXP_BUSI_ARRAY(busi_info_const, each_item) 				= array_info(1)
							EXP_BUSI_ARRAY(busi_monthly_earnings_const, each_item) 	= array_info(2)
							EXP_BUSI_ARRAY(busi_annual_earnings_const, each_item) 	= array_info(3)
							EXP_BUSI_ARRAY(busi_notes_const, each_item) 			= array_info(4)
							known_exp_busi = known_exp_busi + 1
						End If

						If MID(text_line, 7, 14) = "EXP_UNEA_ARRAY" Then
							array_info = Mid(text_line, 24)
							array_info = split(array_info, "~")
							ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, known_exp_unea)

							EXP_UNEA_ARRAY(unea_owner_const, each_item) 			= array_info(0)
							EXP_UNEA_ARRAY(unea_info_const, each_item) 				= array_info(1)
							EXP_UNEA_ARRAY(unea_monthly_earnings_const, each_item) 	= array_info(2)
							EXP_UNEA_ARRAY(unea_weekly_earnings_const, each_item) 	= array_info(3)
							EXP_UNEA_ARRAY(unea_notes_const, each_item) 			= array_info(4)
							known_exp_unea = known_exp_unea + 1
						End If


						If MID(text_line, 7, 14) = "EXP_ACCT_ARRAY" Then
							array_info = Mid(text_line, 24)
							array_info = split(array_info, "~")
							ReDim Preserve EXP_ACCT_ARRAY(account_notes_const, known_exp_acct)

							EXP_ACCT_ARRAY(account_type_const, each_item) 	= array_info(0)
							EXP_ACCT_ARRAY(account_owner_const, each_item) 	= array_info(1)
							EXP_ACCT_ARRAY(bank_name_const, each_item) 		= array_info(2)
							EXP_ACCT_ARRAY(account_amount_const, each_item) = array_info(3)
							EXP_ACCT_ARRAY(account_notes_const, each_item) 	= array_info(4)
							known_exp_acct = known_exp_acct + 1
						End If

					End If
				Next
			End If
		End If
	End With
end function


function verification_dialog()
    If ButtonPressed = verif_button Then
        If second_call <> TRUE Then
            income_source_list = "Select or Type Source"+chr(9)+"Job"+chr(9)+"Self Employment"+chr(9)+"Child Support"+chr(9)+"Social Security Income"+chr(9)+"Unemployment Income"+chr(9)+"VA Income"+chr(9)+"Pension"
            income_verif_time = "[Enter Time Frame]"
            bank_verif_time = "[Enter Time Frame]"
            second_call = TRUE
        End If

        Do
            verif_err_msg = ""

            BeginDialog Dialog1, 0, 0, 610, 385, "Select Verifications"
              Text 280, 10, 120, 10, "Date Verification Request Form Sent:"
              EditBox 400, 5, 50, 15, verif_req_form_sent_date

              GroupBox 530, 35, 75, 145, "PROGRAM(S):"
              Text 535, 48, 65, 40, "Check all programs that require any of the listed verifications:"
              CheckBox 540, 85, 45, 10, "SNAP", verif_snap_checkbox
              CheckBox 540, 95, 45, 10, "CASH", verif_cash_checkbox
              CheckBox 540, 105, 45, 10, "MFIP", verif_mfip_checkbox
              CheckBox 540, 115, 45, 10, "DWP", verif_dwp_checkbox
              CheckBox 540, 125, 45, 10, "MSA", verif_msa_checkbox
              CheckBox 540, 135, 45, 10, "GA", verif_ga_checkbox
              CheckBox 540, 145, 45, 10, "GRH", verif_grh_checkbox
              CheckBox 540, 155, 45, 10, "EMER", verif_emer_checkbox
              CheckBox 540, 165, 45, 10, "HC", verif_hc_checkbox

			  If verif_view = "See All Verifs" Then
			  	Checkbox 60, 45, 200, 10, "Check here to have verifs numbered in the CASE/NOTE.", number_verifs_checkbox
			  	Checkbox 270, 45, 200, 10, "Check here if there are verifs that have been postponed.", verifs_postponed_checkbox

			  	grp_len = 25
				y_pos = 60

				verifs_selected = trim(verifs_selected)
				If right(verifs_selected, 1) = ";" Then
					verifs_to_view = left(verifs_selected, len(verifs_selected)-1)
				Else
					verifs_to_view = verifs_selected
				End If
				If verifs_to_view <> "" Then
					array_of_verifs_selected = ""
					If InStr(verifs_to_view, ";") = 0 Then
						array_of_verifs_needed = array(verifs_to_view)
					Else
						array_of_verifs_needed = split(verifs_to_view, ";")
					End If

					for each verif_item in array_of_verifs_needed
						Text 10, y_pos, 500, 10, verif_item
						y_pos = y_pos + 15
						grp_len = grp_len + 15
					next
				End If
				If y_pos = 60 Then
					Text 10, y_pos, 500, 10, "NO VERIFICATIONS HAVE BEEN LISTED YET"
					grp_len = grp_len + 15
				End If

				GroupBox 5, 35, 520, grp_len, "Verifications Recorded as Requested"
				Text 10, 10, 235, 10, "All verifications you have indicated are listed Here."
				Text 10, 20, 470, 10, "Press 'Add Another' to add other verifications to this list, or add them in the 'ADD VERIFICATION' buttons on the main dialog."
				ButtonGroup ButtonPressed
				  PushButton 485, 10, 50, 15, "Add Another", add_verif_button
			  End If
			  If verif_view = "Add A Verif" Then
	              Groupbox 5, 35, 520, 130, "Personal and Household Information"

	              CheckBox 10, 50, 75, 10, "Verification of ID for ", id_verif_checkbox
	              ComboBox 90, 45, 150, 45, all_the_clients, id_verif_memb
	              CheckBox 300, 50, 100, 10, "Social Security Number for ", ssn_checkbox
	              ComboBox 405, 45, 110, 45, all_the_clients, ssn_verif_memb

	              CheckBox 10, 70, 70, 10, "US Citizenship for ", us_cit_status_checkbox
	              ComboBox 85, 65, 150, 45, all_the_clients, us_cit_verif_memb
	              CheckBox 300, 70, 85, 10, "Immigration Status for", imig_status_checkbox
	              ComboBox 390, 65, 125, 45, all_the_clients, imig_verif_memb

	              CheckBox 10, 90, 90, 10, "Proof of relationship for ", relationship_checkbox
	              ComboBox 105, 85, 150, 45, all_the_clients, relationship_one_verif_memb
	              Text 260, 90, 90, 10, "and"
	              ComboBox 280, 85, 150, 45, all_the_clients, relationship_two_verif_memb

	              CheckBox 10, 110, 85, 10, "Student Information for ", student_info_checkbox
	              ComboBox 100, 105, 150, 45, all_the_clients, student_verif_memb
	              Text 255, 110, 10, 10, "at"
	              EditBox 270, 105, 150, 15, student_verif_source

	              CheckBox 10, 130, 85, 10, "Proof of Pregnancy for", preg_checkbox
	              ComboBox 100, 125, 150, 45, all_the_clients, preg_verif_memb

	              CheckBox 10, 150, 115, 10, "Illness/Incapacity/Disability for", illness_disability_checkbox
	              ComboBox 130, 145, 150, 45, all_the_clients, disa_verif_memb
	              Text 285, 150, 30, 10, "verifying:"
	              EditBox 320, 145, 150, 15, disa_verif_type

                  GroupBox 5, 165, 520, 50, "Income Information"

	              CheckBox 10, 180, 45, 10, "Income for ", income_checkbox
	              ComboBox 60, 175, 140, 45, all_the_clients, income_verif_memb
                  Text 205, 180, 15, 10, "from"
	              ComboBox 225, 175, 125, 45, income_source_list, income_verif_source
                  Text 355, 180, 10, 10, "for"
	              EditBox 370, 175, 145, 15, income_verif_time

	              CheckBox 10, 200, 85, 10, "Employment Status for ", employment_status_checkbox
	              ComboBox 100, 195, 150, 45, all_the_clients, emp_status_verif_memb
	              Text 255, 200, 10, 10, "at"
	              ComboBox 270, 195, 150, 45, employment_source_list, emp_status_verif_source

                  GroupBox 5, 215, 520, 50, "Expense Information"

	              CheckBox 10, 230, 105, 10, "Educational Funds/Costs for", educational_funds_cost_checkbox
	              ComboBox 120, 225, 150, 45, all_the_clients, stin_verif_memb

	              CheckBox 10, 250, 65, 10, "Shelter Costs for ", shelter_checkbox
	              ComboBox 80, 245, 150, 45, all_the_clients, shelter_verif_memb
	              checkBox 240, 250, 175, 10, "Check here if this verif is NOT MANDATORY", shelter_not_mandatory_checkbox

	              GroupBox 5, 265, 600, 30, "Asset Information"

	              CheckBox 10, 280, 70, 10, "Bank Account for", bank_account_checkbox
	              ComboBox 80, 275, 150, 45, all_the_clients, bank_verif_memb
	              Text 235, 280, 45, 10, "account type"
	              ComboBox 285, 275, 145, 45, "Select or Type"+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Certificate of Deposit (CD)"+chr(9)+"Stock"+chr(9)+"Money Market", bank_verif_type
	              Text 435, 280, 10, 10, "for"
	              EditBox 450, 275, 150, 15, bank_verif_time

				  Text 5, 305, 20, 10, "Other:"
				  EditBox 30, 300, 570, 15, other_verifs

				  Text 10, 10, 235, 10, "Check the boxes for any verification you want to add to the CASE/NOTE."
				  Text 10, 25, 490, 10, "Note: After you press 'Update' or 'Return to Dialog' the information from the boxes will be added to the list of verification and the boxes will be 'unchecked'."
				  ButtonGroup ButtonPressed
					PushButton 485, 10, 50, 15, "Update", fill_button
			  End If


              ButtonGroup ButtonPressed
                PushButton 540, 10, 60, 15, "Return to Dialog", return_to_dialog_button
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
                If id_verif_checkbox = checked Then
                    If IsNumeric(left(id_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Identity for Memb " & id_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Identity for " & id_verif_memb & ".; "
                    End If
                    id_verif_checkbox = unchecked
                    id_verif_memb = ""
                End If
                If us_cit_status_checkbox = checked Then
                    If IsNumeric(left(us_cit_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "US Citizenship for Memb " & us_cit_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "US Citizenship for " & us_cit_verif_memb & ".; "
                    End If
                    us_cit_status_checkbox = unchecked
                    us_cit_verif_memb = ""
                End If
                If imig_status_checkbox = checked Then
                    If IsNumeric(left(imig_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Immigration documentation for Memb " & imig_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Immigration documentation for " & imig_verif_memb & ".; "
                    End If
                    imig_status_checkbox = unchecked
                    imig_verif_memb = ""
                End If
                If ssn_checkbox = checked Then
                    If IsNumeric(left(ssn_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Social Security number for Memb " & ssn_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Social Security number for " & ssn_verif_memb & ".; "
                    End If
                    ssn_checkbox = unchecked
                    ssn_verif_memb = ""
                End If
                If relationship_checkbox = checked Then
                    If IsNumeric(left(relationship_one_verif_memb, 2)) = TRUE AND IsNumeric(left(relationship_two_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Relationship between Memb " & relationship_one_verif_memb & " and Memb " & relationship_two_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Relationship between " & relationship_one_verif_memb & " and " & relationship_two_verif_memb & ".; "
                    End If
                    relationship_checkbox = unchecked
                    relationship_one_verif_memb = ""
                    relationship_two_verif_memb = ""
                End If
                If income_checkbox = checked Then
                    If IsNumeric(left(income_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Income for Memb " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    Else
                        verifs_selected = verifs_selected & "Income for " & income_verif_memb & " at " & income_verif_source & " for " & income_verif_time & ".; "
                    End If
                    income_checkbox = unchecked
                    income_verif_source = ""
                    income_verif_memb = ""
                    income_verif_time = ""
                End If
                If employment_status_checkbox = checked Then
                    If IsNumeric(left(emp_status_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Employment Status for Memb " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    Else
                        verifs_selected = verifs_selected & "Employment Status for " & emp_status_verif_memb & " from " & emp_status_verif_source & ".; "
                    End If
                    employment_status_checkbox = unchecked
                    emp_status_verif_memb = ""
                    emp_status_verif_source = ""
                End If
                If student_info_checkbox = checked Then
                    If IsNumeric(left(student_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Student information for Memb " & student_verif_memb & " at " & student_verif_source & ".; "
                    Else
                        verifs_selected = verifs_selected & "Student information for " & student_verif_memb & " at " & student_verif_source & ".; "
                    End If
                    student_info_checkbox = unchecked
                    student_verif_memb = ""
                    student_verif_source = ""
                End If
                If educational_funds_cost_checkbox = checked Then
                    If IsNumeric(left(stin_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Educational funds and costs for Memb " & stin_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Educational funds and costs for " & stin_verif_memb & ".; "
                    End If
                    educational_funds_cost_checkbox = unchecked
                    stin_verif_memb = ""
                End If
                If shelter_checkbox = checked Then
                    If IsNumeric(left(shelter_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Shelter costs for Memb " & shelter_verif_memb & ". "
                    Else
                        verifs_selected = verifs_selected & "Shelter costs for " & shelter_verif_memb & ". "
                    End If
                    If shelter_not_mandatory_checkbox = checked Then verifs_selected = verifs_selected & " THIS VERIFICATION IS NOT MANDATORY."
                    verifs_selected = verifs_selected & "; "
                    shelter_checkbox = unchecked
                    shelter_verif_memb = ""
                End If
                If bank_account_checkbox = checked Then
                    If IsNumeric(left(bank_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & bank_verif_type & " account for Memb " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    Else
                        verifs_selected = verifs_selected & bank_verif_type & " account for " & bank_verif_memb & " for " & bank_verif_time & ".; "
                    End If
                    bank_account_checkbox = unchecked
                    bank_verif_type = ""
                    bank_verif_memb = ""
                    bank_verif_time = ""
                End If
                If preg_checkbox = checked Then
                    If IsNumeric(left(preg_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Pregnancy for Memb " & preg_verif_memb & ".; "
                    Else
                        verifs_selected = verifs_selected & "Pregnancy for " & preg_verif_memb & ".; "
                    End If
                    preg_checkbox = unchecked
                    preg_verif_memb = ""
                End If
                If illness_disability_checkbox = checked Then
                    If IsNumeric(left(disa_verif_memb, 2)) = TRUE Then
                        verifs_selected = verifs_selected & "Ill/Incap or Disability for Memb " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    Else
                        verifs_selected = verifs_selected & "Ill/Incap or Disability for " & disa_verif_memb & " of " & disa_verif_type & ",; "
                    End If
                    illness_disability_checkbox = unchecked
                    disa_verif_memb = ""
                    disa_verif_type = ""
                End If
                other_verifs = trim(other_verifs)
                If other_verifs <> "" Then verifs_selected = verifs_selected & other_verifs & "; "
                other_verifs = ""
            Else
                MsgBox "Additional detail about verifications to note is needed:" & vbNewLine & verif_err_msg
            End If

			If verif_err_msg = "" Then
				If ButtonPressed = add_verif_button Then verif_view = "Add A Verif"
				If ButtonPressed = fill_button Then verif_view = "See All Verifs"
			End If

			If ButtonPressed = add_verif_button Then verif_err_msg = "LOOP" & verif_err_msg
            If ButtonPressed = fill_button Then verif_err_msg = "LOOP" & verif_err_msg
        Loop until verif_err_msg = ""
        ButtonPressed = verif_button
    End If

end function

function write_interview_CASE_NOTE()
	' 'Now we case note!
	STATS_manualtime = STATS_manualtime + 600
	Call start_a_blank_case_note

	CALL write_variable_in_CASE_NOTE("~ Interview Completed on " & interview_date & " ~")
    Call write_bullet_and_variable_in_CASE_NOTE("Case Information", case_summary)
    If cash_request = True and the_process_for_cash = "Application" and type_of_cash = "Family" Then
        Call write_variable_in_CASE_NOTE("Family Cash Program Selection Details")
        CALL write_bullet_and_variable_in_CASE_NOTE("Program selected", family_cash_program)
        CALL write_bullet_and_variable_in_CASE_NOTE("Selection Notes", famliy_cash_notes)
    End If

	CALL write_variable_in_CASE_NOTE("Completed with " & who_are_we_completing_the_interview_with & " via " & how_are_we_completing_the_interview)
	If trim(interpreter_information) <> "" AND interpreter_information <> "No Interpreter Used" Then
		CALL write_variable_in_CASE_NOTE("Interview had interpreter: " & interpreter_information)
		CALL write_variable_in_CASE_NOTE("    Language: " & interpreter_language)
	End If
	If trim(arep_interview_id_information) <> "" Then CALL write_variable_in_CASE_NOTE("AREP Identity Verification: " & arep_interview_id_information)
	If trim(non_applicant_interview_info) <> "" Then CALL write_variable_in_CASE_NOTE("Interviewee Information: " & non_applicant_interview_info)
	CALL write_variable_in_CASE_NOTE("Completed on " & interview_date)
	CALL write_variable_in_CASE_NOTE("Interview using form: " & CAF_form_name & ", received on " & CAF_datestamp)

	CALL write_variable_in_CASE_NOTE("Interview Programs:")

	If cash_request = True Then
		If the_process_for_cash = "Application" Then CALL write_variable_in_CASE_NOTE(" - CASH at Application. App Date: " & CAF_datestamp & ". " & type_of_cash & " Cash.")
		If the_process_for_cash = "Renewal" Then CALL write_variable_in_CASE_NOTE(" - CASH at Renewal. Renewal Month: " & next_cash_revw_mo & "/" & next_cash_revw_yr & ". " & type_of_cash & " Cash.")
		If cash_other_req_detail <> "" Then CALL write_variable_in_CASE_NOTE("   - Request detail: " & cash_other_req_detail)
	End If
	If snap_request = True Then
		If the_process_for_snap = "Application" Then CALL write_variable_in_CASE_NOTE(" - SNAP at Application. App Date: " & CAF_datestamp & ".")
		If the_process_for_snap = "Renewal" Then CALL write_variable_in_CASE_NOTE(" - SNAP at Renewal. Renewal Month: " & next_snap_revw_mo & "/" & next_snap_revw_yr & ".")
		If snap_other_req_detail <> "" Then CALL write_variable_in_CASE_NOTE("   - Request detail: " & snap_other_req_detail)
	End If
	If emer_request = True Then
		CALL write_variable_in_CASE_NOTE(" - EMERGENCY Request at Application. App Date: " & CAF_datestamp & ". EMER is " & type_of_emer)
		If emer_other_req_detail <> "" Then CALL write_variable_in_CASE_NOTE("   - Request detail: " & emer_other_req_detail)
	End If

	CALL write_variable_in_CASE_NOTE("Household Members:")
	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_members) = False Then
            CALL write_variable_in_CASE_NOTE("  * " & HH_MEMB_ARRAY(ref_number, the_members) & "-" & HH_MEMB_ARRAY(full_name_const, the_members))
    		If the_members = 0 Then CALL write_variable_in_CASE_NOTE("    Identity: " & HH_MEMB_ARRAY(id_verif, the_members))
    		If trim(HH_MEMB_ARRAY(client_notes, the_members)) <> "" Then CALL write_variable_in_CASE_NOTE("    NOTES: " & HH_MEMB_ARRAY(client_notes, the_members))
			'REMOVED CLIENT VERIFICATION INFORMATION FOR SUMMARY SCRIPT - it had been here
        End If
	Next
	CALL write_bullet_and_variable_in_CASE_NOTE("Household Info", household_info)
	If edrs_match_found = False Then Call write_variable_in_CASE_NOTE("eDRS run for all Household Members: No DISQ Matches Found")
	If edrs_match_found = True Then
		Call write_variable_in_CASE_NOTE("eDRS run for all Household Members:")
		For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
			If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
                If trim(HH_MEMB_ARRAY(edrs_notes, the_memb)) = "" Then Call write_variable_in_CASE_NOTE("    " & HH_MEMB_ARRAY(edrs_msg, the_memb))
    			If trim(HH_MEMB_ARRAY(edrs_notes, the_memb)) <> "" Then Call write_variable_in_CASE_NOTE("    " & HH_MEMB_ARRAY(edrs_msg, the_memb) & "Notes: " & HH_MEMB_ARRAY(edrs_notes, the_memb))
            End If
		Next
	End If

	CALL write_variable_in_CASE_NOTE("----- ADDR Information -----")
	CALL write_variable_in_CASE_NOTE("Residence Address:")
	CALL write_variable_in_CASE_NOTE("    " & resi_addr_street_full)
	CALL write_variable_in_CASE_NOTE("    " & resi_addr_city & ", " & left(resi_addr_state, 2) & " " & resi_addr_zip)
	CALL write_variable_in_CASE_NOTE("County: " & resi_addr_county)
	If disc_out_of_county = "RESOLVED" Then call write_variable_in_CASE_NOTE("* Household reported living Out of Hennepin County - Case Needs Transfer - additional interview conversation: " & disc_out_of_county_confirmation)
	If trim(reservation_name) = "" Then CALL write_variable_in_CASE_NOTE("    Lives on Reservation: " & reservation_yn)
	If trim(reservation_name) <> "" Then CALL write_variable_in_CASE_NOTE("    Lives on Reservation: " & reservation_yn & " Name: " & reservation_name)
	CALL write_variable_in_CASE_NOTE("    Living Situation: " & living_situation)
	CALL write_variable_in_CASE_NOTE("Reporting Homeless: " & homeless_yn)
	If disc_homeless_no_mail_addr = "RESOLVED" Then call write_variable_in_CASE_NOTE("* Household Experiencing Housing Insecurity - MAIL is Primary Communication of Agency Requests and Actions - additional interview conversation: " & disc_homeless_confirmation)
	If trim(mail_addr_street_full) <> "" OR trim(mail_addr_city) <> "" OR trim(mail_addr_state) <> "" OR trim(mail_addr_zip) <> "" Then
		CALL write_variable_in_CASE_NOTE("Mailing Address:")
		CALL write_variable_in_CASE_NOTE("    " & mail_addr_street_full)
		CALL write_variable_in_CASE_NOTE("    " & mail_addr_city & ", " & left(mail_addr_state, 2) & " " & mail_addr_zip)
	End If
	CALL write_variable_in_CASE_NOTE("Phone Number:")
	If trim(phone_one_number) <> "" Then CALL write_variable_in_CASE_NOTE("    " & phone_one_number & " Type: " & phone_one_type)
	If trim(phone_two_number) <> "" Then CALL write_variable_in_CASE_NOTE("    " & phone_two_number & " Type: " & phone_two_type)
	If trim(phone_three_number) <> "" Then CALL write_variable_in_CASE_NOTE("    " & phone_three_number & " Type: " & phone_three_type)
	If trim(phone_one_number) <> "" AND trim(phone_two_number) <> "" AND trim(phone_three_number) <> "" Then CALL write_variable_in_CASE_NOTE("    No Phone Number provided.")
	If disc_no_phone_number = "RESOLVED" Then call write_variable_in_CASE_NOTE("* No Phone Number was Provided - additional interview conversation: " & disc_phone_confirmation)

	CALL write_variable_in_CASE_NOTE("-----  CAF Information and Notes -----")

	CALL write_bullet_and_variable_in_CASE_NOTE("Earned Income", earned_income_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Unearned Income", unearned_income_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Housing Expenses", housing_expanese_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Other Expenses", other_expenses_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Assets", assets_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes_info)
	CALL write_bullet_and_variable_in_CASE_NOTE("Verbal Changes to " & CAF_form_name, form_changes_info)

	IF create_verif_note = False Then Call write_variable_in_CASE_NOTE("No verifications were indicated at this time.")

	IF create_verif_note = True Then
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE("** VERIFICATIONS REQUESTED - See previous case note for detail")
	End If
	Call write_variable_in_CASE_NOTE("---")
	' RESOURCE NOTIFIER REMOVED FROM HERE

	If qual_questions_yes = FALSE Then Call write_variable_in_CASE_NOTE("* All CAF Qualifying Questions answered 'No'.")

	forms_reviewed = ""
	If rights_responsibilities_checkbox = checked Then forms_reviewed = forms_reviewed & " -4163"
	If privacy_practices_checkbox = checked Then forms_reviewed = forms_reviewed & " -3979"
	If child_support_info_checkbox = checked Then forms_reviewed = forms_reviewed & " -2338 - 3163B"
	If mfip_minor_caregiver_checkbox = checked Then forms_reviewed = forms_reviewed & " -3238"
	If ievs_info_checkbox = checked Then forms_reviewed = forms_reviewed & " -2759"
	If appeal_rights_checkbox = checked or complaints_civil_rights_checkbox = checked Then forms_reviewed = forms_reviewed & " -3353"
	Call write_bullet_and_variable_in_CASE_NOTE("Reviewed DHS Forms", forms_reviewed)

	info_reviewed = ""
	If ebt_card_checkbox = checked Then info_reviewed = info_reviewed & "-EBT Information; "
	If reporting_resp_checkbox = checked Then info_reviewed = info_reviewed & "-Reporting Responsibilities; "
	If program_info_checkbox = checked Then info_reviewed = info_reviewed & "-Program Rules and Eligiblity; "
	If renewal_info_checkbox = checked Then info_reviewed = info_reviewed & "-Information about Renewals; "
	info_reviewed = trim(info_reviewed)
	If right(info_reviewed, 1) = ";" Then info_reviewed = left(info_reviewed, len(info_reviewed)-1)
	Call write_bullet_and_variable_in_CASE_NOTE("Discussed", info_reviewed)

	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

end function


function create_verifs_needed_list(verifs_selected, verifs_needed)
	verifs_needed = verifs_selected
	If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
	If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)

	verifs_needed = trim(verifs_needed)
end function

function write_verification_CASE_NOTE(create_verif_note)
	create_verif_note = False

	Call create_verifs_needed_list(verifs_selected, verifs_needed)

	If trim(verifs_needed) <> "" Then
		create_verif_note = True
	    verif_counter = 1
	    verifs_needed = trim(verifs_needed)
	    If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
	    If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)
	    If InStr(verifs_needed, ";") <> 0 Then
	        verifs_array = split(verifs_needed, ";")
	    Else
	        verifs_array = array(verifs_needed)
	    End If
	End If

    programs_verifs_apply_to = ""
    If verif_snap_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", SNAP"
    If verif_cash_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", CASH"
    If verif_mfip_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", MFIP"
    If verif_dwp_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", DWP"
    If verif_msa_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", MSA"
    If verif_ga_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", GA"
    If verif_grh_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", GRH"
    If verif_emer_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", EMER"
    If verif_hc_checkbox = checked then programs_verifs_apply_to = programs_verifs_apply_to & ", HC"
    If left(programs_verifs_apply_to, 1) = "," Then programs_verifs_apply_to = right(programs_verifs_apply_to, len(programs_verifs_apply_to)-1)
    programs_verifs_apply_to = trim(programs_verifs_apply_to)

	If create_verif_note = True Then

	    Call start_a_blank_CASE_NOTE

	    Call write_variable_in_CASE_NOTE("VERIFICATIONS REQUESTED")

	    Call write_bullet_and_variable_in_CASE_NOTE("Verif request form sent on", verif_req_form_sent_date)

	    Call write_variable_in_CASE_NOTE("---")

	    Call write_variable_in_CASE_NOTE("List of all verifications requested:")
	    If trim(verifs_needed) <> "" Then
		    For each verif_item in verifs_array
		        verif_item = trim(verif_item)
		        If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
		        verif_counter = verif_counter + 1
		        Call write_variable_with_indent_in_CASE_NOTE(verif_item)
				STATS_manualtime = STATS_manualtime + 25
		    Next
		End If
        If programs_verifs_apply_to <> "" Then
            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE("Verifications are needed for " & programs_verifs_apply_to & ".")
        End If
	    If verifs_postponed_checkbox = checked Then
	        Call write_variable_in_CASE_NOTE("---")
	        Call write_variable_in_CASE_NOTE("There may be verifications that are postponed to allow for the approval of Expedited SNAP.")
	    End If
	    Call write_variable_in_CASE_NOTE("---")
	    Call write_variable_in_CASE_NOTE(worker_signature)

	    PF3
	End If


end function


'EXPEDITED DETERMINATION FUNCTIONS------------------------------------------------------------------------------------------------------------------
Function format_explanation_text(text_variable)
	text_variable = trim(text_variable)
	Do while Instr(text_variable, "; ;") <> 0
		text_variable = replace(text_variable, "; ;", "; ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, ";;") <> 0
		text_variable = replace(text_variable, ";;", "; ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, "  ", " ")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, ".; .", "")
		text_variable = trim(text_variable)
	Loop
	Do while Instr(text_variable, "  ") <> 0
		text_variable = replace(text_variable, "; .;", "")
		text_variable = trim(text_variable)
	Loop
	Do while left(text_variable, 1) = "."
		text_variable = right(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
		Do while left(text_variable, 1) = ";"
			text_variable = right(text_variable, len(text_variable) - 1)
			text_variable = trim(text_variable)
		Loop
	Loop
	Do while left(text_variable, 1) = ";"
		text_variable = right(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
	Loop
	Do while right(text_variable, 1) = ";"
		text_variable = left(text_variable, len(text_variable) - 1)
		text_variable = trim(text_variable)
	Loop
	text_variable = trim(text_variable)
End Function

function app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, EXP_JOBS_ARRAY, EXP_BUSI_ARRAY, EXP_UNEA_ARRAY)
	return_btn = 5001
	enter_btn = 5002
	add_another_jobs_btn = 5005
	remove_one_jobs_btn = 5006
	add_another_busi_btn = 5007
	remove_one_busi_btn = 5008
	add_another_unea_btn = 5009
	remove_one_unea_btn = 2010
	income_review_completed = True
	amounts_btn 		= 10

	original_income = determined_income
	determined_income = 0
	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 296, 160, "Determination of Income in Month of Application"
			DropListBox 210, 40, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", jobs_income_yn
			DropListBox 210, 60, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", busi_income_yn
			DropListBox 235, 110, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", unea_income_yn
			ButtonGroup ButtonPressed
				PushButton 240, 140, 50, 15, "Enter", enter_btn
			Text 10, 10, 205, 10, "Does this household have any income?"
			GroupBox 10, 25, 255, 65, "Earned Income "
			Text 65, 45, 140, 10, "Is anyone in the household working a job?"
			Text 25, 65, 180, 10, "Does anyone in the household have self employment?"
			GroupBox 10, 95, 280, 40, "Unearned Income"
			Text 20, 115, 215, 10, "Does anyone in the household receive any other kind of income?"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			income_review_completed = False
			Exit Do
		End If

		If jobs_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from a Job."
		If busi_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from Self Employment."
		If unea_income_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has Income from Another Source."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If income_review_completed = True Then
		Do
			prvt_err_msg = ""

			If jobs_income_yn = "No" Then jobs_grp_len = 30
			If jobs_income_yn = "Yes" Then jobs_grp_len = 55 + (UBound(EXP_JOBS_ARRAY, 2) + 1) * 20
			If busi_income_yn = "No" Then busi_grp_len = 30
			If busi_income_yn = "Yes" Then busi_grp_len = 55 + (UBound(EXP_BUSI_ARRAY, 2) + 1) * 20
			If unea_income_yn = "No" Then unea_grp_len = 30
			If unea_income_yn = "Yes" Then unea_grp_len = 55 + (UBound(EXP_UNEA_ARRAY, 2) + 1) * 20

			'determining if additional length of the dialog is needed to display interview notes about income from the main script
			interview_note_details_exists = False
			intvw_notes_len = 20
			If trim(question_8_interview_notes) <> "" Then
				interview_note_details_exists = True
				intvw_notes_len = intvw_notes_len + 20
			End If
			If trim(question_10_interview_notes) <> "" Then
				interview_note_details_exists = True
				intvw_notes_len = intvw_notes_len + 20
			End If
			If trim(question_12_interview_notes) <> "" Then
				interview_note_details_exists = True
				intvw_notes_len = intvw_notes_len + 20
			End If

			dlg_len = 45 + jobs_grp_len + busi_grp_len + unea_grp_len
			If interview_note_details_exists = True Then dlg_len = dlg_len + intvw_notes_len + 10

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, dlg_len, "Determination of Income in Month of Application"
			  	ButtonGroup ButtonPressed
					'displaying details from interview notes in the dialog for calculating app month income
				  	y_pos = 10
					GroupBox 10, y_pos, 380, jobs_grp_len, "JOBS"
					y_pos = y_pos + 15
					If jobs_income_yn = "Yes" Then
						Text 20, y_pos, 190, 10, "JOBS Income on this case"
						y_pos = y_pos + 15
						Text 20, y_pos, 50, 10, "Employee"
						Text 90, y_pos, 70, 10, "Employer/Job"
						Text 185, y_pos, 50, 10, "Hourly Wage"
						Text 245, y_pos, 50, 10, "Weekly Hours"
						Text 305, y_pos, 50, 10, "Pay Frequency"
						y_pos = y_pos + 10

						For the_job = 0 to UBound(EXP_JOBS_ARRAY, 2)
							EXP_JOBS_ARRAY(jobs_wage_const, the_job) = EXP_JOBS_ARRAY(jobs_wage_const, the_job) & ""
							EXP_JOBS_ARRAY(jobs_hours_const, the_job) = EXP_JOBS_ARRAY(jobs_hours_const, the_job) & ""
							EditBox 20, y_pos, 60, 15, EXP_JOBS_ARRAY(jobs_employee_const, the_job)
							EditBox 90, y_pos, 85, 15, EXP_JOBS_ARRAY(jobs_employer_const, the_job)
							EditBox 185, y_pos, 50, 15, EXP_JOBS_ARRAY(jobs_wage_const, the_job)
							EditBox 245, y_pos, 50, 15, EXP_JOBS_ARRAY(jobs_hours_const, the_job)
							DropListBox 305, y_pos, 75, 15, "Select One..."+chr(9)+"Weekly"+chr(9)+"Biweekly"+chr(9)+"Semi-Monthly"+chr(9)+"Monthly", EXP_JOBS_ARRAY(jobs_frequency_const, the_job)
							y_pos = y_pos + 20
						Next
						PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_jobs_btn
						PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_jobs_btn
						y_pos = y_pos + 20
					Else
						Text 20, y_pos, 355, 10, "This household does NOT have JOBS."
						y_pos = y_pos + 20
					End If

					GroupBox 10, y_pos, 380, busi_grp_len, "Self Employment"
					y_pos = y_pos + 15
					If busi_income_yn = "Yes" Then
						Text 20, y_pos, 190, 10, "BUSI Income on this case"
						y_pos = y_pos + 15
						Text 20, y_pos, 65, 10, "Business Owner"
						Text 125, y_pos, 70, 10, "Business"
						Text 230, y_pos, 65, 10, "Monthly Earnings"
						Text 290, y_pos, 65, 10, "Annual Earnings"
						y_pos = y_pos + 10
						' Text 305, y_pos, 50, 10, "Pay Frequency"
						For the_busi = 0 to UBound(EXP_BUSI_ARRAY, 2)
							EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) & ""
							EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) = EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) & ""

							EditBox 20, y_pos, 95, 15, EXP_BUSI_ARRAY(busi_owner_const, the_busi)
							EditBox 125, y_pos, 95, 15, EXP_BUSI_ARRAY(busi_info_const, the_busi)
							EditBox 230, y_pos, 50, 15, EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)
							EditBox 290, y_pos, 50, 15, EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi)
							y_pos = y_pos + 20
						Next
						PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_busi_btn
						PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_busi_btn
						y_pos = y_pos + 20
					Else
						Text 20, y_pos, 355, 10, "This household does NOT have BUSI."
						y_pos = y_pos + 20
					End If

					GroupBox 10, y_pos, 380, unea_grp_len, "Unearned"
					y_pos = y_pos + 15
					If unea_income_yn = "Yes" Then
						Text 20, y_pos, 190, 10, "UNEA Income on this case"
						y_pos = y_pos + 15
						Text 20, y_pos, 65, 10, "Member Receiving"
						Text 125, y_pos, 70, 10, "Income Type"
						Text 230, y_pos, 65, 10, "Monthly Amount"
						Text 290, y_pos, 65, 10, "Weekly Amount"
						y_pos = y_pos + 10
						' Text 305, y_pos, 50, 10, "Pay Frequency"
						For the_unea = 0 to UBound(EXP_UNEA_ARRAY, 2)
							EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea) & ""
							EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) & ""
							EditBox 20, y_pos, 95, 15, EXP_UNEA_ARRAY(unea_owner_const, the_unea)
							EditBox 125, y_pos, 95, 15, EXP_UNEA_ARRAY(unea_info_const, the_unea)
							EditBox 230, y_pos, 50, 15, EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea)
							EditBox 290, y_pos, 50, 15, EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea)
							y_pos = y_pos + 20
						Next
						PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_unea_btn
						PushButton 320, y_pos, 60, 10, "REMOVE ONE", remove_one_unea_btn
						y_pos = y_pos + 20
					Else
						Text 20, y_pos, 355, 10, "This household does NOT have UNEA."
						y_pos = y_pos + 20
					End If

					PushButton 345, dlg_len - 20, 50, 15, "Return", return_btn
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then
				income_review_completed = False
				Exit Do
			End If

			last_jobs_item = UBound(EXP_JOBS_ARRAY, 2)
			If ButtonPressed = add_another_jobs_btn Then
				last_jobs_item = last_jobs_item + 1
				ReDim Preserve EXP_JOBS_ARRAY(jobs_notes_const, last_jobs_item)
			End If
			If ButtonPressed = remove_one_jobs_btn Then
				last_jobs_item = last_jobs_item - 1
				ReDim Preserve EXP_JOBS_ARRAY(jobs_notes_const, last_jobs_item)
			End If

			last_busi_item = UBound(EXP_BUSI_ARRAY, 2)
			If ButtonPressed = add_another_busi_btn Then
				last_busi_item = last_busi_item + 1
				ReDim Preserve EXP_BUSI_ARRAY(busi_notes_const, last_busi_item)
			End If
			If ButtonPressed = remove_one_unea_btn Then
				last_busi_item = last_busi_item - 1
				ReDim Preserve EXP_BUSI_ARRAY(busi_notes_const, last_busi_item)
			End If

			last_unea_item = UBound(EXP_UNEA_ARRAY, 2)
			If ButtonPressed = add_another_unea_btn Then
				last_unea_item = last_unea_item + 1
				ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, last_unea_item)
			End If
			If ButtonPressed = remove_one_busi_btn Then
				last_unea_item = last_unea_item - 1
				ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, last_unea_item)
			End If
			If ButtonPressed = -1 Then ButtonPressed = return_btn

            If jobs_income_yn = "Yes" Then
    			For the_job = 0 to UBound(EXP_JOBS_ARRAY, 2)
    				EXP_JOBS_ARRAY(jobs_employee_const, the_job) = trim(EXP_JOBS_ARRAY(jobs_employee_const, the_job))
    				EXP_JOBS_ARRAY(jobs_employer_const, the_job) = trim(EXP_JOBS_ARRAY(jobs_employer_const, the_job))
    				EXP_JOBS_ARRAY(jobs_wage_const, the_job) = trim(EXP_JOBS_ARRAY(jobs_wage_const, the_job))
    				EXP_JOBS_ARRAY(jobs_hours_const, the_job) = trim(EXP_JOBS_ARRAY(jobs_hours_const, the_job))
    				EXP_JOBS_ARRAY(jobs_frequency_const, the_job) = trim(EXP_JOBS_ARRAY(jobs_frequency_const, the_job))

    				If EXP_JOBS_ARRAY(jobs_employee_const, the_job) <> "" OR EXP_JOBS_ARRAY(jobs_employer_const, the_job) <> "" OR EXP_JOBS_ARRAY(jobs_wage_const, the_job) <> "" OR EXP_JOBS_ARRAY(jobs_hours_const, the_job) <> "" Then
    					jobs_err_msg = ""
    					If EXP_JOBS_ARRAY(jobs_employee_const, the_job) = "" Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the name of the employer for this JOB."
    					If EXP_JOBS_ARRAY(jobs_employer_const, the_job) = "" Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the employer for This JOB."
    					If IsNumeric(EXP_JOBS_ARRAY(jobs_wage_const, the_job)) = False Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the amount that " & EXP_JOBS_ARRAY(jobs_employee_const, the_job) & " is paid per hour from " & EXP_JOBS_ARRAY(jobs_employer_const, the_job) & " as a number."
    					If IsNumeric(EXP_JOBS_ARRAY(jobs_hours_const, the_job)) = False Then jobs_err_msg = jobs_err_msg & vbCr & "* Enter the number of hours " & EXP_JOBS_ARRAY(jobs_employee_const, the_job) & " works per week in the application month for " & EXP_JOBS_ARRAY(jobs_employer_const, the_job) & " as a number."
    					If EXP_JOBS_ARRAY(jobs_frequency_const, the_job) = "Select One..." Then jobs_err_msg = jobs_err_msg & vbCr & "* Select the pay frequency that " & EXP_JOBS_ARRAY(jobs_employee_const, the_job) & " receives their checks in from " & EXP_JOBS_ARRAY(jobs_employer_const, the_job) & "."
    					If jobs_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the JOB that is Number " & the_job + 1 & " on the list." & vbCr & jobs_err_msg & vbCr
    				End If
    			Next
            End If

            If busi_income_yn = "Yes" Then
    			For the_busi = 0 to UBound(EXP_BUSI_ARRAY, 2)
    				EXP_BUSI_ARRAY(busi_owner_const, the_busi) = trim(EXP_BUSI_ARRAY(busi_owner_const, the_busi))
    				EXP_BUSI_ARRAY(busi_info_const, the_busi) = trim(EXP_BUSI_ARRAY(busi_info_const, the_busi))
    				EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = trim(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi))
    				EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) = trim(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi))

    				If EXP_BUSI_ARRAY(busi_owner_const, the_busi) <> "" OR EXP_BUSI_ARRAY(busi_info_const, the_busi) <> "" OR EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) <> "" OR EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) <> "" Then
    					busi_err_msg = ""
    					If EXP_BUSI_ARRAY(busi_owner_const, the_busi) = "" Then busi_err_msg = busi_err_msg & vbCr & "* Enter the name of the employer for this Self Employment."
    					If EXP_BUSI_ARRAY(busi_info_const, the_busi) = "" Then busi_err_msg = busi_err_msg & vbCr & "* Enter the business information for this Self Employment."
    					If EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) <> "" AND IsNumeric(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = False Then busi_err_msg = busi_err_msg & vbCr & "* Enter the amount that " & EXP_BUSI_ARRAY(busi_owner_const, the_busi) & " earns monthly from " & EXP_BUSI_ARRAY(busi_info_const, the_busi) & "."
    					If EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) <> "" AND IsNumeric(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = False Then busi_err_msg = busi_err_msg & vbCr & "* Enter the number of hours " & EXP_BUSI_ARRAY(busi_owner_const, the_busi) & " earns yearly from " & EXP_BUSI_ARRAY(busi_info_const, the_busi) & "."
    					If IsNumeric(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True AND IsNumeric(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = True Then
    						EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = FormatNumber(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi), 2, -1, 0, -1)
    						EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) = FormatNumber(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi), 2, -1, 0, -1)
    						annual_from_monthly = 0
    						annual_from_monthly =  EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) * 12
    						annual_from_monthly = FormatNumber(annual_from_monthly, 2, -1, 0, -1)
    						If annual_from_monthly <> EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) Then busi_err_msg = busi_err_msg & vbCr & "* The annual amount does not match up with the monthly amount entered. The Annual earnings should be 12 times the Monthly earnings. You only need to enter one of these amounts."
    					ElseIf IsNumeric(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True AND EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) = "" Then
    						EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi) = FormatNumber(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)*12, 2, -1, 0, -1)
    					ElseIf IsNumeric(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi)) = True AND EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = "" Then
    						EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi) = FormatNumber(EXP_BUSI_ARRAY(busi_annual_earnings_const, the_busi)/12, 2, -1, 0, -1)
    					End If
    					If busi_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the BUSI that is Number " & the_busi + 1 & " on the list." & vbCr & busi_err_msg & vbCr
    				End If
    			Next
            End If

            If unea_income_yn = "Yes" Then
    			For the_unea = 0 to UBound(EXP_UNEA_ARRAY, 2)
    				unea_err_msg = ""
    				EXP_UNEA_ARRAY(unea_owner_const, the_unea) = trim(EXP_UNEA_ARRAY(unea_owner_const, the_unea))
    				EXP_UNEA_ARRAY(unea_info_const, the_unea) = trim(EXP_UNEA_ARRAY(unea_info_const, the_unea))
    				EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = trim(EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea))
    				EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = trim(EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea))
    				If EXP_UNEA_ARRAY(unea_owner_const, the_unea) <> "" OR EXP_UNEA_ARRAY(unea_info_const, the_unea) <> "" OR EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea) <> "" OR EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) <> "" Then
    					If EXP_UNEA_ARRAY(unea_owner_const, the_unea) = "" Then unea_err_msg = unea_err_msg & vbCr & "* Enter the name of the the person who received this Unearned Income."
    					If EXP_UNEA_ARRAY(unea_info_const, the_unea) = "" Then unea_err_msg = unea_err_msg & vbCr & "* Enter the information of what type of Unearned Income this is listed."
    					If IsNumeric(EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = True AND IsNumeric(EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = True Then
    						If FormatNumber(EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea), 0) <> FormatNumber(EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) * 4.3, 0) Then unea_err_msg = unea_err_msg & vbCr & "* Enter Only one of the following: Weekly Amount or Monthly Amount"
    					ElseIf IsNumeric(EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = False AND EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) = "" Then
    						unea_err_msg = unea_err_msg & vbCr & "* Enter the amount that " & EXP_UNEA_ARRAY(unea_owner_const, the_unea) & " receives per month from " & EXP_UNEA_ARRAY(unea_info_const, the_unea) & " as a number."
    					ElseIf IsNumeric(EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = False AND EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea) = "" Then
    						unea_err_msg = unea_err_msg & vbCr & "* Enter the number of hours " & EXP_UNEA_ARRAY(unea_owner_const, the_unea) & " receives per week from " & EXP_UNEA_ARRAY(unea_info_const, the_unea) & " as a number."
    					End IF
    					If unea_err_msg <> "" Then prvt_err_msg = prvt_err_msg & vbCr & "For the UNEA that is Number " & the_unea + 1 & " on the list." & vbCr & unea_err_msg & vbCr
    				End If
    			Next
            End If

			If prvt_err_msg <> "" AND ButtonPressed = return_btn Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop Until ButtonPressed = return_btn AND prvt_err_msg = ""
	End If

	For the_job = 0 to UBound(EXP_JOBS_ARRAY, 2)
		If IsNumeric(EXP_JOBS_ARRAY(jobs_wage_const, the_job)) = True AND IsNumeric(EXP_JOBS_ARRAY(jobs_hours_const, the_job)) = True Then
			weekly_pay = EXP_JOBS_ARRAY(jobs_wage_const, the_job) * EXP_JOBS_ARRAY(jobs_hours_const, the_job)
			EXP_JOBS_ARRAY(jobs_monthly_pay_const, the_job) = weekly_pay * 4.3
			determined_income = determined_income + EXP_JOBS_ARRAY(jobs_monthly_pay_const, the_job)
		End If
	Next

	For the_busi = 0 to UBound(EXP_BUSI_ARRAY, 2)
		If IsNumeric(EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)) = True Then determined_income = determined_income + EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi)
	Next
	For the_unea = 0 to UBound(EXP_UNEA_ARRAY, 2)
		If IsNumeric(EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea)) = True Then
			determined_income = determined_income + EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea)
		ElseIf IsNumeric(EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea)) = True Then
			monthly_pay = EXP_UNEA_ARRAY(unea_weekly_earnings_const, the_unea) * 4.3
			determined_income = determined_income + monthly_pay
		End If
	Next
	determined_income = FormatNumber(determined_income, 2, -1, 0, -1)

	If income_review_completed = False Then determined_income = original_income

	determined_income = determined_income & ""
	ButtonPressed = amounts_btn
end function

function app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, EXP_ACCT_ARRAY)
	return_btn = 5001
	enter_btn = 5002
	add_another_btn = 5003
	remove_one_btn = 5004
	amounts_btn 		= 10

	assets_review_completed = True

	original_assets = determined_assets
	determined_assets = 0
	If cash_amount_yn <> "Yes" OR bank_account_yn <> "Yes" Then
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 271, 135, "Determination of Assets in Month of Application"
			  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  GroupBox 10, 25, 255, 40, "Cash"
			  Text 25, 45, 155, 10, "Does the household have any Cash Savings?"
			  DropListBox 180, 40, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", cash_amount_yn
			  GroupBox 10, 70, 255, 40, "Accounts"
			  Text 20, 90, 190, 10, "Does anyone in the household have any Bank Accounts?"
			  DropListBox 210, 85, 45, 45, "?"+chr(9)+"Yes"+chr(9)+"No", bank_account_yn
			  ButtonGroup ButtonPressed
			    PushButton 215, 115, 50, 15, "Enter", enter_btn
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then
				assets_review_completed = False
				Exit Do
			End If

			If cash_amount_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has CASH."
			If bank_account_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter if the household has A BANK ACCOUNT."

			If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop until prvt_err_msg = ""
	End If

	If assets_review_completed = True Then
		Do
			prvt_err_msg = ""
			cash_amount = cash_amount & ""

			If cash_amount_yn = "No" Then cash_grp_len = 30
			If cash_amount_yn = "Yes" Then cash_grp_len = 50
			If bank_account_yn = "No" Then acct_grp_len = 30
			If bank_account_yn = "Yes" Then acct_grp_len = 60 + (UBound(EXP_ACCT_ARRAY, 2) + 1) * 20

			interview_note_details_exists = False
			intvw_notes_len = 20

			If trim(question_20_interview_notes) <> "" Then
				interview_note_details_exists = True
				intvw_notes_len = intvw_notes_len + 30
			End If

			dlg_len = 55 + cash_grp_len + acct_grp_len
			If interview_note_details_exists = True Then dlg_len = dlg_len + intvw_notes_len + 10


			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 351, dlg_len, "Determination of Assets in Month of Application"
				y_pos = 10

				'displaying details from interview notes in the dialog for calculating app month assets
				If interview_note_details_exists = True Then
					GroupBox 10, y_pos, 335, intvw_notes_len, "Interview NOTES entered into the Script already"
					y_pos = y_pos + 15
					If trim(question_20_interview_notes) <> "" Then
						Text 20, y_pos, 320, 10, "20. Does anyone in the household have assets?"
						Text 25, y_pos+10, 315, 20, question_20_interview_notes
						y_pos = y_pos + 30
					End If
					y_pos = y_pos + 10

				End If

				Text 10, y_pos, 205, 10, "Are there any Liquid Assets available to the household?"
				y_pos = y_pos + 15
				GroupBox 10, y_pos, 220, cash_grp_len, "Cash"
				y_pos = y_pos + 15
				If cash_amount_yn = "Yes" Then
					Text 20, y_pos, 155, 10, "This household HAS Cash Savings."
					y_pos = y_pos + 15
					Text 20, y_pos, 150, 10, "How much in total does the household have?"
					EditBox 175, y_pos - 5, 45, 15, cash_amount
					y_pos = y_pos + 25
				Else
					Text 20, y_pos, 155, 10, "This household does NOT have Cash."
					y_pos = y_pos + 20
				End If
				GroupBox 10, y_pos, 335, acct_grp_len, "Accounts"
				y_pos = y_pos + 15
				If bank_account_yn = "Yes" Then
					Text 20, y_pos, 190, 10, "This household HAS Bank Accounts."
					y_pos = y_pos + 15
					Text 20, y_pos, 50, 10, "Account Type"
					Text 90, y_pos, 70, 10, "Owner of Account"
					Text 180, y_pos, 45, 10, "Bank Name"
					Text 285, y_pos, 35, 10, "Amount"
					y_pos = y_pos + 15

					For the_acct = 0 to UBound(EXP_ACCT_ARRAY, 2)
						EXP_ACCT_ARRAY(account_amount_const, the_acct) = EXP_ACCT_ARRAY(account_amount_const, the_acct) & ""
						DropListBox 20, y_pos, 60, 45, "Select One..."+chr(9)+"Checking"+chr(9)+"Savings"+chr(9)+"Other", EXP_ACCT_ARRAY(account_type_const, the_acct)
						EditBox 90, y_pos, 85, 15, EXP_ACCT_ARRAY(account_owner_const, the_acct)
						EditBox 180, y_pos, 100, 15, EXP_ACCT_ARRAY(bank_name_const, the_acct)
						EditBox 285, y_pos, 50, 15, EXP_ACCT_ARRAY(account_amount_const, the_acct)
						y_pos = y_pos + 20
					Next
				Else
					Text 20, y_pos, 155, 10, "This household does NOT have Bank Accounts."
				End If
				ButtonGroup ButtonPressed
					If bank_account_yn = "Yes" Then PushButton 20, y_pos, 60, 10, "ADD ANOTHER", add_another_btn
					If bank_account_yn = "Yes" Then PushButton 275, y_pos, 60, 10, "REMOVE ONE", remove_one_btn
					PushButton 295, dlg_len - 20, 50, 15, "Return", return_btn
			EndDialog

			dialog Dialog1
			If ButtonPressed = 0 Then
				assets_review_completed = False
				Exit Do
			End If

			last_acct_item = UBound(EXP_ACCT_ARRAY, 2)
			If ButtonPressed = add_another_btn Then
				last_acct_item = last_acct_item + 1
				ReDim Preserve EXP_ACCT_ARRAY(account_notes_const, last_acct_item)
			End If
			If ButtonPressed = remove_one_btn Then
				last_acct_item = last_acct_item - 1
				ReDim Preserve EXP_ACCT_ARRAY(account_notes_const, last_acct_item)
			End If
			If ButtonPressed = -1 Then ButtonPressed = return_btn

			cash_amount = trim(cash_amount)
			If cash_amount <> "" And IsNumeric(cash_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Cash Amount as a number."

			For the_acct = 0 to UBound(EXP_ACCT_ARRAY, 2)
				EXP_ACCT_ARRAY(account_amount_const, the_acct) = trim(EXP_ACCT_ARRAY(account_amount_const, the_acct))
				If EXP_ACCT_ARRAY(account_amount_const, the_acct) <> "" And IsNumeric(EXP_ACCT_ARRAY(account_amount_const, the_acct)) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the Bank Account amounts as a member."
				If EXP_ACCT_ARRAY(account_type_const, the_acct)	= "Select One..." Then prvt_err_msg = prvt_err_msg & vbCr & "* Select the Bank Account type."
			Next
			If prvt_err_msg <> "" AND ButtonPressed = return_btn Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
		Loop Until ButtonPressed = return_btn AND prvt_err_msg = ""

		If cash_amount = "" Then cash_amount = 0
		cash_amount = cash_amount * 1
		For the_acct = 0 to UBound(EXP_ACCT_ARRAY, 2)
			If EXP_ACCT_ARRAY(account_amount_const, the_acct) = "" Then EXP_ACCT_ARRAY(account_amount_const, the_acct) = 0
			EXP_ACCT_ARRAY(account_amount_const, the_acct) = EXP_ACCT_ARRAY(account_amount_const, the_acct) * 1
			determined_assets = determined_assets + EXP_ACCT_ARRAY(account_amount_const, the_acct)
		Next
		determined_assets = determined_assets + cash_amount
	End If
	If assets_review_completed = False Then determined_assets =  original_assets

	determined_assets = determined_assets & ""
	ButtonPressed = amounts_btn
end function

function app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
	return_btn = 5001
	amounts_btn 		= 10

	shel_review_completed = True
	rent_amount = rent_amount & ""
	lot_rent_amount = lot_rent_amount & ""
	mortgage_amount = mortgage_amount & ""
	insurance_amount = insurance_amount & ""
	tax_amount = tax_amount & ""
	room_amount = room_amount & ""
	garage_amount = garage_amount & ""
	subsidy_amount = subsidy_amount & ""

	original_shel = determined_shel
	determined_shel = 0

	dlg_len = 140

	interview_note_details_exists = False
	intvw_notes_len = 20

	If trim(question_14_interview_notes) <> "" Then
		interview_note_details_exists = True
		intvw_notes_len = intvw_notes_len + 40
	End If

	If interview_note_details_exists = True Then dlg_len = dlg_len + intvw_notes_len + 5

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 196, dlg_len, "Determination of Housing Cost in Month of Application"
			EditBox 45, 35, 50, 15, rent_amount
			EditBox 45, 55, 50, 15, lot_rent_amount
			EditBox 45, 75, 50, 15, mortgage_amount
			EditBox 45, 95, 50, 15, insurance_amount
			EditBox 140, 35, 50, 15, tax_amount
			EditBox 140, 55, 50, 15, room_amount
			EditBox 140, 75, 50, 15, garage_amount
			EditBox 140, 95, 50, 15, subsidy_amount
			Text 10, 15, 165, 10, "Enter the total Shelter Expense for the Houshold."
			Text 25, 40, 20, 10, "Rent:"
			Text 10, 60, 35, 10, " Lot Rent:"
			Text 10, 80, 35, 10, "Mortgage:"
			Text 10, 100, 35, 10, "Insurance:"
			Text 115, 40, 25, 10, "Taxes:"
			Text 115, 60, 25, 10, "Room:"
			Text 110, 80, 30, 10, "Garage:"
			Text 105, 100, 35, 10, "  Subsidy:"

			y_pos = 120
			'displaying details from interview notes in the dialog for calculating app month housing expenses
			If interview_note_details_exists = True Then
				GroupBox 5, y_pos, 185, intvw_notes_len, "Interview NOTES entered into the Script already"
				y_pos = y_pos + 15
				If trim(question_14_interview_notes) <> "" Then
					Text 10, y_pos, 175, 10, "14. Does your household have housing expenses?"
					Text 15, y_pos+10, 170, 30, question_14_interview_notes
					y_pos = y_pos + 40
				End If
				y_pos = y_pos + 10

			End If
			ButtonGroup ButtonPressed
				PushButton 140, y_pos, 50, 15, "Return", return_btn

		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			shel_review_completed = False
			Exit Do
		End If

		rent_amount = trim(rent_amount)
		lot_rent_amount = trim(lot_rent_amount)
		mortgage_amount = trim(mortgage_amount)
		insurance_amount = trim(insurance_amount)
		tax_amount = trim(tax_amount)
		room_amount = trim(room_amount)
		garage_amount = trim(garage_amount)
		subsidy_amount = trim(subsidy_amount)

		If rent_amount <> "" AND IsNumeric(rent_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the RENT amount as a number."
		If lot_rent_amount <> "" AND IsNumeric(lot_rent_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the LOT RENT amount as a number."
		If mortgage_amount <> "" AND IsNumeric(mortgage_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the MORTGAGE amount as a number."
		If insurance_amount <> "" AND IsNumeric(insurance_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the INSURANCE amount as a number."
		If tax_amount <> "" AND IsNumeric(tax_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the TAXES amount as a number."
		If room_amount <> "" AND IsNumeric(room_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the ROOM amount as a number."
		If garage_amount <> "" AND IsNumeric(garage_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the GARAGE amount as a number."
		If subsidy_amount <> "" AND IsNumeric(subsidy_amount) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the SUBSIDY amount as a number."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If IsNumeric(rent_amount) = True Then determined_shel = determined_shel + rent_amount
	If IsNumeric(lot_rent_amount) = True Then determined_shel = determined_shel + lot_rent_amount
	If IsNumeric(mortgage_amount) = True Then determined_shel = determined_shel + mortgage_amount
	If IsNumeric(insurance_amount) = True Then determined_shel = determined_shel + insurance_amount
	If IsNumeric(tax_amount) = True Then determined_shel = determined_shel + tax_amount
	If IsNumeric(room_amount) = True Then determined_shel = determined_shel + room_amount
	If IsNumeric(garage_amount) = True Then determined_shel = determined_shel + garage_amount

	If shel_review_completed = False Then determined_shel = original_shel

	determined_shel = determined_shel & ""
	ButtonPressed = amounts_btn
end function

function app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)
	calculate_btn = 5000
	return_btn = 5001
	amounts_btn 		= 10
	determined_utilities = 0
	If heat_expense = True then heat_checkbox = checked
	If ac_expense = True then ac_checkbox = checked
	If electric_expense = True then electric_checkbox = checked
	If phone_expense = True then phone_checkbox = checked
	If none_expense = True then none_checkbox = checked

	dlg_len = 175

	interview_note_details_exists = False
	intvw_notes_len = 20

	If trim(question_15_interview_notes) <> "" Then
		interview_note_details_exists = True
		intvw_notes_len = intvw_notes_len + 40
	End If

	If interview_note_details_exists = True Then dlg_len = dlg_len + intvw_notes_len + 10

	Do
		current_utilities = all_utilities

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 246, dlg_len, "Determination of Utilities in Month of Application"
			CheckBox 30, 45, 50, 10, "Heat", heat_checkbox
			CheckBox 30, 60, 65, 10, "Air Conditioning", ac_checkbox
			CheckBox 30, 75, 50, 10, "Electric", electric_checkbox
			CheckBox 30, 90, 50, 10, "Phone", phone_checkbox
			CheckBox 30, 105, 50, 10, "NONE", none_checkbox

			Text 10, 10, 235, 10, "Check the boxes for each utility the household is responsible to pay:"
			GroupBox 15, 30, 225, 95, "Utilities"
			Text 150, 45, 50, 10, "$ " & determined_utilities
			Text 150, 60, 35, 35, all_utilities
			Text 15, 135, 225, 20, "Remember, this expense could be shared, they are still considered responsible to pay and we count the WHOLE standard."

			y_pos = 160
			'displaying details from interview notes in the dialog for calculating app month utilities
			If interview_note_details_exists = True Then
				GroupBox 5, y_pos, 235, intvw_notes_len, "Interview NOTES entered into the Script already"
				y_pos = y_pos + 15
				If trim(question_15_interview_notes) <> "" Then
					Text 10, y_pos, 215, 10, "15. Does your household have utility expenses any time during the year?"
					Text 15, y_pos+10, 210, 30, question_15_interview_notes
					y_pos = y_pos + 40
				End If
				y_pos = y_pos + 10
			Else
				y_pos = y_pos - 5
			End If

			ButtonGroup ButtonPressed
				PushButton 170, 105, 65, 15, "Calculate", calculate_btn
				PushButton 170, y_pos, 65, 15, "Return", return_btn
		EndDialog

		dialog Dialog1

		some_vs_none_discrepancy = False
		If (heat_checkbox = checked OR ac_checkbox = checked OR electric_checkbox = checked OR phone_checkbox = checked) AND none_checkbox = checked Then some_vs_none_discrepancy = True
		If some_vs_none_discrepancy = True Then MsgBox "Attention:" & vbCr & vbCr & "You have selected NONE and selected at least one other utility expense. If it is NONE, then no other utilities should be checked."

		all_utilities = ""
		If heat_checkbox = checked Then all_utilities = all_utilities & ", Heat"
		If ac_checkbox = checked Then all_utilities = all_utilities & ", AC"
		If electric_checkbox = checked Then all_utilities = all_utilities & ", Electric"
		If phone_checkbox = checked Then all_utilities = all_utilities & ", Phone"
		If none_checkbox = checked Then all_utilities = all_utilities & ", None"
		If left(all_utilities, 2) = ", " Then all_utilities = right(all_utilities, len(all_utilities) - 2)

		If all_utilities = current_utilities AND ButtonPressed = -1 Then ButtonPressed = return_btn

		determined_utilities = 0
		If heat_checkbox = checked OR ac_checkbox = checked Then
			determined_utilities = determined_utilities + heat_AC_amt
		Else
			If electric_checkbox = checked Then determined_utilities = determined_utilities + electric_amt
			If phone_checkbox = checked Then determined_utilities = determined_utilities + phone_amt
		End If

	Loop Until ButtonPressed = return_btn And some_vs_none_discrepancy = False

	heat_expense = False
	ac_expense = False
	electric_expense = False
	phone_expense = False
	none_expense = False

	If heat_checkbox = checked Then heat_expense = True
	If ac_checkbox = checked Then ac_expense = True
	If electric_checkbox = checked Then electric_expense = True
	If phone_checkbox = checked Then phone_expense = True
	If none_checkbox = checked Then none_expense = True

	ButtonPressed = amounts_btn
end function

Function determine_actions(case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four, is_elig_XFS, snap_denial_date, approval_date, CAF_datestamp, do_we_have_applicant_id, action_due_to_out_of_state_benefits, mn_elig_begin_date, other_snap_state, case_has_previously_postponed_verifs_that_prevent_exp_snap, delay_action_due_to_faci, deny_snap_due_to_faci)

	case_assesment_text = ""
	next_steps_one = ""
	next_steps_two = ""
	next_steps_three = ""
	next_steps_four = ""
	If IsDate(snap_denial_date) = True Then
		case_assesment_text = "DENIAL has been determined - Case does not meet 'All Other Eligibility Criteria'."
		next_steps_one = "Complete the DENIAL by updating MAXIS and enter a full, detaild DENIAL CASE/NOTE. Complete the full processing before moving on to your next task."

		If action_due_to_out_of_state_benefits = "DENY" Then
			add_msg = "Update MEMI with out of state benefit information to generate accurate DENIAL Results. Add a WCOM to the denial advising resident to reapply within 30 days of the benefits ending in the other state."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If
		End If
		If deny_snap_due_to_faci = True Then
			add_msg = "Ensure FACI is coded correctly for accurate DENIAL. Add a WCOM to the denials advising resident to rapply when release from the facility is within 30 days."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If
		End If

		add_msg = "Process this denial quickly as a PENDING SNAP case will continue to be assigned until acted on, once the determination is done and action can be taken, we do not want to reassign this case."
		If next_steps_two = "" then
			next_steps_two = add_msg
		ElseIf next_steps_three = "" Then
			next_steps_three = add_msg
		ElseIf next_steps_four = "" Then
			next_steps_four = add_msg
		End If

		add_msg = "Denials can be coded in REPT/PND2 if they are for a resident 'Withdraw' of their request. Otherwise, since the interview should be done at this point, denials should be processed in STAT."
		If next_steps_three = "" Then
			next_steps_three = add_msg
		ElseIf next_steps_four = "" Then
			next_steps_four = add_msg
		End If

		add_msg = "It is best practice to add detail to the Denial WCOM for clarity for the resident."
		If next_steps_four = "" Then next_steps_four = add_msg
	ElseIf is_elig_XFS = True Then
		If IsDate(approval_date) = True Then
			case_assesment_text = "Case IS EXPEDITED and ready to approve"
			next_steps_one = "Approve SNAP Expedited package of " & expedited_package & " before moving on to your next task. Update MAXIS STAT panels to generate EXPEDITED SNAP Eligibility Results and APPROVE."

			If action_due_to_out_of_state_benefits = "APPROVE" AND mn_elig_begin_date <> CAF_datestamp Then
				If DateDiff("d", date, mn_elig_begin_date) > 0 Then
					add_msg = "After approval, send a BENE request in SIR to have benefits issued on " & mn_elig_begin_date & " instead of the regular issuance day."
					If next_steps_two = "" then
						next_steps_two = add_msg
					ElseIf next_steps_three = "" Then
						next_steps_three = add_msg
					ElseIf next_steps_four = "" Then
						next_steps_four = add_msg
					End If
				End If
			End If

			add_msg = "Remember, EXPEDITED is based on income, assets, and shelter/utility expenses only. Even having a delay reason does not mean the case is not still EXPEDITED."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "We attempt to approve expedited within 24 hours of the date of application, or as close to that time as possible. It is crucial we complete the approval at the time we determine the case to be EXPEDITED."
			If next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "EBT Card information can be found below, but often requires contact with the resident, remember REI issuances can prevent residents from receiving their card."
			If next_steps_four = "" Then next_steps_four = add_msg
		Else
			case_assesment_text = "Case IS EXPEDITED but approval must be delayed."
			next_steps_one = "We must strive to approve this case for the EXPEDITED package of " & expedited_package & " as soon as possible. Make every effort to complete the requirements of this delay and approve the case"

			If do_we_have_applicant_id = False Then
				add_msg = "Double check the case file for ANY document that can be used as an identity document.Advise resident to get us ANY form of ID they can, MNbenefits or the virtual dropbox may be quickest way to receive this document."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If
			If action_due_to_out_of_state_benefits = "FOLLOW UP" Then
				If other_snap_state <> "" Then add_msg = "Contact " & other_snap_state & " as soon as possible to determine the end date of of SNAP in " & other_snap_state & "."
				If other_snap_state = "" Then add_msg = "Contact the other state as soon as possible to determine the end date of of SNAP in that state."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
				add_msg = "This case needs regular review to be able to approve SNAP as soon as, the current verifications come in OR the previous verifications come in. Assist the resident in getting any verifications that we can."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			If delay_action_due_to_faci = True Then
				add_msg = "Advise resident and the facility to contact us as soon as possible to be able to approve SNAP once the resident leaves the facility."
				If next_steps_two = "" then
					next_steps_two = add_msg
				ElseIf next_steps_three = "" Then
					next_steps_three = add_msg
				ElseIf next_steps_four = "" Then
					next_steps_four = add_msg
				End If
			End If

			add_msg = "Delays in processing Expedited should be few and far between, we must make every reasonable effort to get these cases approved as quickly as possible."
			If next_steps_two = "" then
				next_steps_two = add_msg
			ElseIf next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "Check in with Knowledge Now about this case, as delays cause negative impact on our timeliness reports."
			If next_steps_three = "" Then
				next_steps_three = add_msg
			ElseIf next_steps_four = "" Then
				next_steps_four = add_msg
			End If

			add_msg = "Remember, EXPEDITED is based on income, assets, and shelter/utility expenses only. Even having a delay reason does not mean the case is not still EXPEDITED."
			If next_steps_four = "" Then next_steps_four = add_msg
		End If
	ElseIf is_elig_XFS = False Then
		case_assesment_text = "Case is NOT EXPEDITED, approval decision should follow standard SNAP Policy."
		next_steps_one = "If there are mandatory verifications, request them immediately. If all verifications have been received, process the case right away."
		next_steps_two = ""
		next_steps_three = ""
		next_steps_four = ""
	End If
end function

function determine_calculations(determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)
	determined_income = trim(determined_income)
	If determined_income = "" Then determined_income = 0
	determined_income = determined_income * 1

	determined_assets = trim(determined_assets)
	If determined_assets = "" Then determined_assets = 0
	determined_assets = determined_assets * 1

	determined_shel = trim(determined_shel)
	If determined_shel = "" Then determined_shel = 0
	determined_shel = determined_shel * 1

	determined_utilities = trim(determined_utilities)
	If determined_utilities = "" Then determined_utilities = 0
	determined_utilities = determined_utilities * 1

	calculated_resources = determined_income + determined_assets
	calculated_expenses = determined_shel + determined_utilities

	calculated_low_income_asset_test = False
	calculated_resources_less_than_expenses_test = False
	is_elig_XFS = False

	If determined_income < 150 AND determined_assets <= 100 Then calculated_low_income_asset_test = True
	If calculated_resources < calculated_expenses Then calculated_resources_less_than_expenses_test = True

	If calculated_low_income_asset_test = True OR calculated_resources_less_than_expenses_test = True Then is_elig_XFS = True

	determined_income = determined_income & ""
	determined_assets = determined_assets & ""
	determined_shel = determined_shel & ""
	determined_utilities = determined_utilities & ""
end function

function snap_in_another_state_detail(CAF_datestamp, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)
	original_snap_denial_date = snap_denial_date
	original_snap_denial_reason = snap_denial_explain
	calculation_done = False
	other_state_benefits_openended = False
	action_due_to_out_of_state_benefits = ""
	' other_snap_state = "MN - Minnesota"
	day_30_from_application = DateAdd("d", 30, CAF_datestamp)
	calculate_btn = 5000
	return_btn = 5001
	determination_btn = 20

	Do
		Do
			prvt_err_msg = ""

			Dialog1 = ""
			If calculation_done = False Then BeginDialog Dialog1, 0, 0, 381, 190, "Case Received SNAP in Another State"
			If calculation_done = True Then BeginDialog Dialog1, 0, 0, 381, 295, "Case Received SNAP in Another State"
			  DropListBox 255, 55, 110, 45, "Select One..."+chr(9)+state_list, other_snap_state
			  EditBox 255, 75, 60, 15, other_state_reported_benefit_end_date
			  CheckBox 40, 95, 320, 10, "Check here if resident reports the benefits are NOT ended or it is UKNOWN if they are ended.", other_state_benefits_not_ended_checkbox
			  DropListBox 255, 110, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", other_state_contact_yn
			  EditBox 255, 130, 60, 15, other_state_verified_benefit_end_date
			  ButtonGroup ButtonPressed
			    PushButton 325, 170, 50, 15, "Calculate", calculate_btn
			  Text 10, 10, 365, 10, "If a Household has received SNAP in another state, we may still be able to issue Expedited SNAP in Minnesota. "
			  Text 10, 25, 320, 10, "Complete the following information to get guidance on handling cases with SNAP in another State:"
			  GroupBox 10, 45, 365, 120, "Other State Benefits"
			  Text 20, 60, 235, 10, "What State is the Household / Resident receiving SNAP benefits from?"
			  Text 40, 80, 215, 10, "When is the resident REPORTING benefits ending in this state?"
			  Text 20, 115, 230, 10, "Have you called the other state to confirm / discover the SNAP status?"
			  Text 20, 135, 230, 10, "What end date has been confirmed / verified for the other state SNAP?"

			  If calculation_done = True Then
				  GroupBox 10, 190, 365, 80, "Resolution"
				  If action_due_to_out_of_state_benefits = "DENY" Then
					  Text 20, 205, 205, 20, "SNAP should be denied as the other state end date is AFTER the 30 day processing period of the application in MN."
					  Text 245, 205, 120, 10, "Date of Application: " & CAF_datestamp
					  If IsDate(other_state_verified_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_verified_benefit_end_date
					  ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
					  	Text 255, 215, 110, 10, "End Of Benefits: " & other_state_reported_benefit_end_date
					  End If
					  ' Text 30, 230, 120, 10, "SNAP Denial Date: " & snap_denial_date
					  ' Text 30, 240, 335, 30, "Denial Reason: " & snap_denial_explain
				  ElseIf action_due_to_out_of_state_benefits = "APPROVE" Then
					  Text 20, 205, 205, 20, "SNAP should be APPROVED "
					  Text 245, 205, 120, 10, "Date of Application: " & CAF_datestamp
					  Text 25, 215, 175, 10, "Eligibility can start in MN as of " & mn_elig_begin_date
					  If other_state_contact_yn <> "Yes" Then
					  	Text 20, 230, 340, 10, "Verification of out of state eligibility end can be postponed "
						Text 20, 240, 340, 10, "We should make reasonable efforts to obtain verification so, "
						Text 20, 250, 340, 10, "it is best to attempt a call to the other state right away for verification."
					  End If
				  ElseIf action_due_to_out_of_state_benefits = "FOLLOW UP" Then
					  Text 20, 205, 205, 20, "You must connect with the other state to determine when the benefits have ended or IF the benefits will end."
				  End If
				  ButtonGroup ButtonPressed
				    PushButton 325, 275, 50, 15, "Return", return_btn
			  End If
			EndDialog

			dialog Dialog1

			If ButtonPressed = 0 Then Exit Do

			If IsDate(other_state_reported_benefit_end_date) = False AND other_state_benefits_not_ended_checkbox = unchecked Then prvt_err_msg = prvt_err_msg & vbCr & "* We cannot complete the calculation if a reported end date has not been entered."
			If IsDate(other_state_reported_benefit_end_date) = True AND other_state_benefits_not_ended_checkbox = checked Then prvt_err_msg = prvt_err_msg & vbCr & "* You have entered an end date AND indicated the benefits have not ended by checking the box. Please select only one."

			If IsDate(other_state_reported_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_reported_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the REPORTED end date."
			End If
			If IsDate(other_state_verified_benefit_end_date) = True Then
				If DatePart("d", DateAdd("d", 1, other_state_verified_benefit_end_date)) <> 1 Then prvt_err_msg = prvt_err_msg & vbCr & "* SNAP Eligiblity end dates should be the last day of the month that the household received SNAP benefits for. Update the date to be the LAST day of the last month of eligiblity in the other state for the VERIFIED end date."
			End If
			If prvt_err_msg <> "" Then
				MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
				calculation_done = False
			End If

		Loop until prvt_err_msg = ""

		If ButtonPressed = 0 Then
			calculation_done = False
			Exit Do
		End If

		calculation_done = True
		If other_snap_state = "NB - MN Newborn" OR other_snap_state = "MN - Minnesota" OR other_snap_state = "Select One..." OR other_snap_state = "FC - Foreign Country" OR other_snap_state = "UN - Unknown" Then other_snap_state = ""
		If IsDate(other_state_verified_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_verified_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_verified_benefit_end_date)
			End If
		ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_reported_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_reported_benefit_end_date)
			End If
		ElseIf other_state_benefits_not_ended_checkbox = checked Then
			action_due_to_out_of_state_benefits = "FOLLOW UP"
			other_state_benefits_openended = True
		End If
		If action_due_to_out_of_state_benefits <> "DENY" Then
			snap_denial_date = original_snap_denial_date
			snap_denial_explain = original_snap_denial_reason
		End If
		If action_due_to_out_of_state_benefits <> "APPROVE" Then expedited_package = original_expedited_package
	Loop until ButtonPressed = return_btn
	If action_due_to_out_of_state_benefits = "APPROVE" Then
		If DateDiff("d", mn_elig_begin_date, CAF_datestamp) > 0 Then
			mn_elig_begin_date = CAF_datestamp
			expedited_package = original_expedited_package
		Else
			MN_elig_month = DatePart("m", mn_elig_begin_date)
			MN_elig_month = right("0"&MN_elig_month, 2)
			MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
			expedited_package = MN_elig_month & "/" & MN_elig_year
		End If
	End If
	If action_due_to_out_of_state_benefits = "DENY" Then
		snap_denial_date = date
		If other_snap_state = "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days"
		If other_snap_state <> "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days"
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_explain = snap_denial_explain & "; " & deny_msg & "."
	End If
	If action_due_to_out_of_state_benefits <> "DENY" Then
		If other_snap_state = "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in other state. Household can reapply once the eligibility in another state is ending within 30 days"
		If other_snap_state <> "" Then deny_msg = "Active SNAP in another state exists past the end of the 30 day application processing window. There is no eligibility in MN until the benefits have ended in " & other_snap_state & ". Household can reapply once the eligibility in another state is ending within 30 days"
		snap_denial_explain = replace(snap_denial_explain, deny_msg, "")
	End If
	snap_denial_date = snap_denial_date & ""
	ButtonPressed = determination_btn
end function

function previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_CAF_datestamp, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
	fn_review_btn = 5005
	return_btn = 5001
	determination_btn = 20
	prev_post_verif_assessment_done = True
	case_has_previously_postponed_verifs_that_prevent_exp_snap = False

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 446, 160, "Case Previously Received EXP SNAP with Postponed Verifications"
		  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
		  Text 10, 30, 125, 10, "This does not apply to cases where:"
		  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
		  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
		  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
		  Text 15, 85, 175, 15, "What is the DATE OF APPLICATION for the Expedited Approval that had Postponed Verifications?"
		  EditBox 195, 85, 50, 15, previous_CAF_datestamp
		  Text 275, 110, 115, 10, "Are these verifications mandatory?"
		  DropListBox 400, 105, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_verifs_mandatory_yn
		  Text 15, 110, 175, 10, "List the verifications that were previously postponed:"
		  EditBox 15, 120, 425, 15, prev_verif_list
		  Text 15, 145, 220, 10, "Does the case have Postponed Verifications for THIS Application?"
		  DropListBox 235, 140, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", curr_verifs_postponed_yn
		  ButtonGroup ButtonPressed
		    PushButton 390, 140, 50, 15, "Review", fn_review_btn
		EndDialog

		dialog Dialog1

		If ButtonPressed = 0 Then
			prev_post_verif_assessment_done = False
			Exit Do
		End If

		prev_verif_list = trim(prev_verif_list)
		If IsDate(previous_CAF_datestamp) = False Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the date of application from the last time this case received an Expedited SNAP approval WITH Postponed Verifications."
		If prev_verifs_mandatory_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* You must review the verifications that were previously postponed and enter them here."
		If prev_verif_list = "" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the verifications that were previously postponed and indicate if any of them were mandatory."
		If curr_verifs_postponed_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Indicate if the CURRENT application has verifications required that would need to be postponed to approve the Expedited SNAP."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If prev_post_verif_assessment_done = True Then
		PREVIOUS_footer_month = DatePart("m", previous_CAF_datestamp)
		PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

		PREVIOUS_footer_year = right(DatePart("yyyy", previous_CAF_datestamp), 2)

		If DatePart("d", previous_CAF_datestamp) > 15 Then
			second_month_of_previous_exp_package = DateAdd("m", 1, previous_CAF_datestamp)
			PREVIOUS_footer_month = DatePart("m", second_month_of_previous_exp_package)
			PREVIOUS_footer_month = right("0"&PREVIOUS_footer_month, 2)

			PREVIOUS_footer_year = right(DatePart("yyyy", second_month_of_previous_exp_package), 2)
		End If
		previous_expedited_package = PREVIOUS_footer_month & "/" & PREVIOUS_footer_year

		ask_more_questions = False
		If IsDate(previous_CAF_datestamp) = True AND prev_verifs_mandatory_yn = "Yes" AND curr_verifs_postponed_yn = "Yes" Then ask_more_questions = True
		If ask_more_questions = True Then
			Do
				prvt_err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 436, 110, "Case Previously Received EXP SNAP with Postponed Verifications"
				  Text 10, 10, 435, 10, "A case that was approved Expedited SNAP with postponed verifications MAY not be able to have Expedited Approved right away."
				  Text 10, 30, 125, 10, "This does not apply to cases where:"
				  Text 15, 40, 165, 10, "- The Postponed Verification were not mandatory."
				  Text 15, 50, 275, 10, "- The Postponed Verification were provided - even if Eligibility was not approved."
				  Text 15, 60, 385, 10, "- The case met all criteria for Regular SNAP to be issued and was approved for 'Ongoing' SNAP for at least one month."
				  Text 10, 80, 180, 10, "Did the case get approved for any SNAP after " & previous_expedited_package & "?"
				  DropListBox 195, 75, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", ongoing_snap_approved_yn
				  Text 20, 95, 170, 10, "Check ECF, are the postponed verifications on file?"
				  DropListBox 195, 90, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", prev_post_verifs_recvd_yn
				  ButtonGroup ButtonPressed
				    PushButton 380, 90, 50, 15, "Review", fn_review_btn

				  Text 10, 270, 280, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
				  Text 15, 295, 210, 10, "- Provide all verifications that were postponed and mandatory."
				  Text 15, 305, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
				  Text 20, 315, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
				EndDialog

				dialog Dialog1

				If ButtonPressed = 0 Then
					prev_post_verif_assessment_done = False
					Exit Do
				End If

				If ongoing_snap_approved_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review MAXIS and determine if SNAP was approved after the last month of the expedited package (" & previous_expedited_package & "). If it was, the case met all requirements to gain SNAP eligibility."
				If prev_post_verifs_recvd_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Review the ECF case file and see if the mandatory postponed verifications were ever received, even if SNAP was not approved."

				If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
			Loop until prvt_err_msg = ""
		End If
	End If

	If prev_post_verif_assessment_done = True Then
		If ask_more_questions = False OR ongoing_snap_approved_yn = "Yes" OR prev_post_verifs_recvd_yn = "Yes" Then
			Dialog1 = ""
			y_pos = 85

			BeginDialog Dialog1, 0, 0, 436, 120, "Case Previously Received EXP SNAP with Postponed Verifications"
			  GroupBox 10, 10, 415, 55, "EXPEDITED CAN BE APPROVED"
			  Text 25, 25, 100, 10, "Based on this case situation"
			  Text 30, 35, 325, 10, "This case CAN be approved for Expedited without a delay due to Previous Postponed Verifications."
			  Text 35, 45, 285, 10, "(There may be another reason for delay, complete the rest of the review to determine.)"
			  Text 15, 75, 45, 10, "Explanation:"
			  If prev_verifs_mandatory_yn = "No" Then
				  Text 15, y_pos, 350, 10, "The previously postponed verifications were not mandatory, so case met all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If curr_verifs_postponed_yn = "No" Then
				  Text 15, y_pos, 350, 10, "There are no verifications that are required and being postponed now, so case meets all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If ongoing_snap_approved_yn = "Yes" Then
				  Text 15, y_pos, 350, 10, "Case was approved regular SNAP after the expedited package time, so case met all SNAP eligibility criteria."
				  y_pos = y_pos + 10
			  End If
			  If prev_post_verifs_recvd_yn = "Yes" Then
				  Text 50, y_pos, 350, 10, "The postponed verifications have been received, which meets the requirement to receive another posponed verification approval package."
				  y_pos = y_pos + 10
			  End If
			  ButtonGroup ButtonPressed
			    PushButton 380, 100, 50, 15, "Update", update_btn
			EndDialog

			dialog Dialog1

		End If

		If ask_more_questions = True AND ongoing_snap_approved_yn = "No" AND prev_post_verifs_recvd_yn = "No" Then
			case_has_previously_postponed_verifs_that_prevent_exp_snap = True

			BeginDialog Dialog1, 0, 0, 291, 145, "Case Previously Received EXP SNAP with Postponed Verifications"
			  GroupBox 5, 5, 280, 60, "EXPEDITED APPROVAL MUST BE DELAYED"
			  Text 20, 20, 100, 10, "Based on this case situation"
			  Text 25, 30, 195, 10, "This case CANNOT be approved for Expedited at this time."
			  Text 30, 40, 235, 20, "The case would require postponing verifications when we already have allowed for postponed verifications that have not been received."
			  Text 10, 70, 275, 20, "If a case cannot be approved due to previously not received Postponed Verifications, the case must meet ONE of the following criteria:"
			  Text 15, 95, 210, 10, "- Provide all verifications that were postponed and mandatory."
			  Text 15, 105, 280, 10, "- Meet all criterea to approve SNAP - including receipt of all mandatory verifications."
			  Text 20, 115, 265, 20, "(This means if a case has no verifications to request, we CAN approve Expedited as the case meets all criteria to approve SNAP.)"
			  ButtonGroup ButtonPressed
			    PushButton 235, 125, 50, 15, "Update", update_btn
			EndDialog

			dialog Dialog1
		End If
	End If
	If prev_post_verif_assessment_done = False Then
		case_has_previously_postponed_verifs_that_prevent_exp_snap = False
		Explain_not_completed_msg = Msgbox("All of the details around postponed verifications have not been entered to be able to determine if there should be a delay due to previously postponed verifications." & vbCr & vbCr & "If you have details to record and you wish to complete the assesment, press the button for this functionality again and the script will restart the questions.", vbOK, "Escape Pressed - Details not Completed")
	End If
	delay_msg = "Approval cannot be completed as case has postponed verifications when postpone verifications were previously allowed and not provided, nor has the case meet 'ongoing SNAP' eligibility"
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = False Then delay_explanation = replace(delay_explanation, delay_msg, "")
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
		If InStr(delay_explanation, delay_msg) = 0 Then delay_explanation = delay_explanation & "; " & delay_msg & "."
	End If

	ButtonPressed = determination_btn
end function

function household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
	return_btn = 5001
	determination_btn = 20
	delay_action_due_to_faci = False
	deny_snap_due_to_faci = False
	faci_review_completed = True

	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 266, 200, "Case Previously Received EXP SNAP with Postponed Verifications"
		  EditBox 70, 40, 180, 15, facility_name
		  DropListBox 210, 60, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", snap_inelig_faci_yn
		  EditBox 110, 100, 50, 15, faci_entry_date
		  EditBox 110, 120, 50, 15, faci_release_date
		  CheckBox 110, 140, 150, 10, "Check here if the release date is unknown.", release_date_unknown_checkbox
		  DropListBox 210, 155, 40, 45, "?"+chr(9)+"Yes"+chr(9)+"No", release_within_30_days_yn
		  ButtonGroup ButtonPressed
		    PushButton 215, 180, 45, 15, "Return", return_btn
		  Text 10, 10, 90, 10, "Resident is in a Facility"
		  GroupBox 10, 25, 250, 55, "Facility Information"
		  Text 20, 45, 50, 10, "Facility Name"
		  Text 95, 65, 115, 10, "Is this a 'SNAP Ineligible' facility?"
		  GroupBox 10, 85, 250, 90, "Resident Stay Information"
		  Text 20, 105, 85, 10, "Date of Entry into Facility:"
		  Text 30, 125, 75, 10, "Date of Exit / Release:"
		  Text 165, 125, 45, 10, "(or expected)"
		  Text 20, 160, 185, 10, "Does the resident expect to be released by " & day_30_from_application & "?"
		EndDialog

		dialog Dialog1
		If ButtonPressed = 0 Then
			faci_review_completed = False
			Exit Do
		End If

		facility_name = trim(facility_name)
		If facility_name = "" Then prvt_err_msg = prvt_err_msg & vbCr & "* Enter the name of the facility."
		If snap_inelig_faci_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Select if this is a SNAP Ineligible Facility."
		If IsDate(faci_release_date) = False AND release_date_unknown_checkbox = unchecked Then prvt_err_msg = prvt_err_msg & vbCr & "* Either enter a release date (expected release date) or indicate that the release date is unknown."
		If IsDate(faci_release_date) = True AND release_date_unknown_checkbox = checked Then prvt_err_msg = prvt_err_msg & vbCr & "* You have entered a release date AND indicated the release date is unknown."
		If release_date_unknown_checkbox = checked AND release_within_30_days_yn = "?" Then prvt_err_msg = prvt_err_msg & vbCr & "* Since the expected release date is unknown, indicate if this release is expected to be prior to do the end of the 30 day processing period."

		If prvt_err_msg <> "" Then MsgBox "***** Additional Action/Information Needed ******" & vbCr & vbCr & "Please resolve to continue:" & vbCr & prvt_err_msg
	Loop until prvt_err_msg = ""

	If faci_review_completed = True Then
		If snap_inelig_faci_yn = "Yes" Then
			If IsDate(faci_release_date) = True Then
				If DateDiff("d", date, faci_release_date) > 0 AND DateDiff("d", faci_release_date, day_30_from_application) >= 0 Then delay_action_due_to_faci = True
				If DateDiff("d", date, faci_release_date) > 0 AND DateDiff("d", faci_release_date, day_30_from_application) < 0 Then deny_snap_due_to_faci = True
			ElseIf release_date_unknown_checkbox = checked Then
				If release_within_30_days_yn = "Yes" Then delay_action_due_to_faci = True
				If release_within_30_days_yn = "No" Then deny_snap_due_to_faci = True
 			End If
		End If
	End If

	delay_msg = "Approval cannot be completed as resident is still in an Ineligible SNAP Facility"
	If delay_action_due_to_faci = False Then delay_explanation = replace(delay_explanation, delay_msg, "")
	If delay_action_due_to_faci = True Then
		If InStr(delay_explanation, delay_msg) = 0 Then delay_explanation = delay_explanation & "; " & delay_msg & "."
	End If

	deny_msg = "SNAP to be denied as resident is in an Ineligible SNAP Facility and is not expected to be released within 30 days of the Date of Application"
	If deny_snap_due_to_faci = False Then
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_date = ""
		snap_denial_explain = replace(snap_denial_explain, deny_msg, "")
	End If
	If deny_snap_due_to_faci = True Then
		If InStr(snap_denial_explain, deny_msg) = 0 Then snap_denial_explain = snap_denial_explain & "; " & deny_msg & "."
		snap_denial_date = date
		snap_denial_date = snap_denial_date & ""
	End If

	If faci_review_completed = True Then
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 216, 130, "Case Previously Received EXP SNAP with Postponed Verifications"
		  Text 10, 10, 90, 10, "Resident is in a Facility"
		  ButtonGroup ButtonPressed
		    PushButton 165, 110, 45, 15, "Return", return_btn
		  Text 15, 25, 140, 20, "The resident's stay in the Facility impacts the SNAP Expedited Processing by:"
		  If delay_action_due_to_faci = True Then Text 20, 55, 195, 10, "Delaying the Approval of Expedited until the Release Date"
		  If deny_snap_due_to_faci = True Then Text 20, 55, 190, 20, "The SNAP case should be DENIED as the resident will not be released within 30 days."
		  If delay_action_due_to_faci = False AND deny_snap_due_to_faci = False Then Text 20, 55, 195, 10, "No change to the Expedited processing because:"
		  y_pos = 65
		  If snap_inelig_faci_yn = "No" Then
			  Text 30, y_pos, 180, 10, "The Facility is not a SNAP Ineligible Facility."
			  y_pos = y_pos + 10
		  End If
		  If IsDate(faci_release_date) = True Then
			  If DateDiff("d", date, faci_release_date) <= 0 Then
			  	Text 30, y_pos, 180, 30, "The release date has already happend. SNAP Eligibility Begin date should be changed to " & faci_release_date & " and processed based on the rest of the case information."
			  End If
		  End If
		EndDialog

		dialog Dialog1
	End If

	ButtonPressed = determination_btn
end function

function send_support_email_to_KN()

	email_subject = "Assistance with Case at SNAP Application - Possible EXP"
	If developer_mode = True Then email_subject = "TESTING RUN - " & email_subject & " - can be deleted"

	email_body = "I am completing a SNAP Expedited Determination." & vbCr & vbCr
	email_body = email_body & "Case Number: " & MAXIS_case_number & vbCr & vbCr
	email_body = email_body & "Amounts currently entered at the Determination:" & vbCr
	email_body = email_body & "Income: $ " & determined_income & vbCr
	email_body = email_body & "Assets: $ " & determined_assets & vbCr
	email_body = email_body & "Housing: $ " & determined_shel & vbCr
	email_body = email_body & "Utilities: $ " & determined_utilities & vbCr & vbCr
	email_body = email_body & "Script Calculations:" & vbCr
	If is_elig_XFS = True Then email_body = email_body & "Case IS EXPEDITED." & vbCr
	If is_elig_XFS = False Then email_body = email_body & "Case is NOT Expedtied." & vbCr
	email_body = email_body & "Unit has less than $150 monthly Gross Income AND $100 or less in assets: " & calculated_low_income_asset_test & vbCr
	email_body = email_body & "Unit's combined resources are less than housing expense: " & calculated_resources_less_than_expenses_test & vbCr & vbCr
	email_body = email_body & "Case Dates/Timelines:" & vbCr
	email_body = email_body & "Date of Application: " & CAF_datestamp & vbCr
	email_body = email_body & "Date of Interview: " & interview_date & vbCr
	email_body = email_body & "Date of Approval: " & approval_date & " (or planned date of approval)" & vbCr
	email_body = email_body & "Processing Delay Explanation: " & delay_explanation & vbCr
	email_body = email_body & "SNAP Denial Date: " & snap_denial_date & vbCr
	email_body = email_body & "Denial Explanation: " & snap_denial_explain & vbCr & vbCr
	email_body = email_body & "Other Information:" & vbCr
	If applicant_id_on_file_yn <> "" AND applicant_id_on_file_yn <> "?" Then email_body = email_body & "Is there an ID on file for the applicant? " & applicant_id_on_file_yn & vbCr
	If applicant_id_through_SOLQ <> "" AND applicant_id_through_SOLQ <> "?" Then email_body = email_body & "Can the Identity of the applicant be cleard through SOLQ/SMI? " & applicant_id_through_SOLQ & vbCr
	If postponed_verifs_yn <> "" AND postponed_verifs_yn <> "?" Then email_body = email_body & "Are there Postponed Verifications for this case? " & postponed_verifs_yn & vbCr
	If trim(list_postponed_verifs) <> "" Then email_body = email_body & "Postponed Verifications: " & list_postponed_verifs & vbCr
	If action_due_to_out_of_state_benefits <> "" Then
		email_body = email_body & "Other SNAP State: " & other_snap_state & vbCr
		email_body = email_body & "Reported End Date: " & other_state_reported_benefit_end_date & vbCr
		If other_state_benefits_openended = True Then email_body = email_body & "End date of SNAP in other state not determined." & vbCr
		email_body = email_body & "Has other State End Date been Confirmed/Verified: " & other_state_contact_yn & vbCr
		email_body = email_body & "Verified End Date: " & other_state_verified_benefit_end_date & vbCr
		email_body = email_body & "Action recommended by script based on information provided: " & action_due_to_out_of_state_benefits & vbCr
	End If
	If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then email_body = email_body & "It appears this case has postponed verifications from a previous EXP SNAP package that prevent approval of a new Expedited Package." & vbCr & vbCr

	email_body = email_body & "---" & vbCr
	If worker_name <> "" Then email_body = email_body & "Signed, " & vbCr & worker_name

	email_body = "~~This email is generated from wihtin the 'Expedited Determination' Script.~~" & vbCr & vbCr & email_body
	Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
	'Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", False)
	'create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
end function
'---------------------------------------------------------------------------------------------------------------------------


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'THIS IS WHERE CONSTANTS AND DIMMING HAPPENS IN THE INTERVIEW SCRIPT - REMOVED FROM HERE BECAUSE WE CANNOT REDEFINE CONSTANTS

Dim household_info, earned_income_info, unearned_income_info, housing_expanese_info
Dim other_expenses_info, assets_info, other_notes_info, form_changes_info
Dim rights_responsibilities_checkbox, privacy_practices_checkbox, ebt_card_checkbox, complaints_civil_rights_checkbox
Dim reporting_resp_checkbox, program_info_checkbox, child_support_info_checkbox, mfip_minor_caregiver_checkbox
Dim renewal_info_checkbox, ievs_info_checkbox, appeal_rights_checkbox

Dim show_case_info


show_pg_one_memb01_and_exp	= 1
show_pg_one_address			= 2
show_pg_memb_list			= 3
show_case_info				= 4
show_qual					= 5
show_pg_last				= 6
show_arep_page				= 7
expedited_determination		= 8


show_exp_pg_amounts = 1
show_exp_pg_determination = 2
show_exp_pg_review = 3

update_addr = FALSE
update_pers = FALSE
page_display = 1
discrepancies_exist = False
children_under_18_in_hh = False
children_under_22_in_hh = False
school_age_children_in_hh = False
expedited_determination_needed = False
expedited_determination_completed = False

intv_exp_pay_heat_checkbox = unchecked
intv_exp_pay_ac_checkbox = unchecked
intv_exp_pay_electricity_checkbox = unchecked
intv_exp_pay_phone_checkbox = unchecked
intv_exp_pay_none_checkbox = unchecked
qual_question_one = "?"
qual_question_two = "?"
qual_question_three = "?"
qual_question_four = "?"
qual_question_five = "?"
disc_no_phone_number = "N/A"
disc_homeless_no_mail_addr = "N/A"
disc_out_of_county = "N/A"
disc_rent_amounts = "N/A"
disc_utility_amounts = "N/A"
disc_yes_phone_no_expense = "N/A"
disc_no_phone_yes_expense = "N/A"
verif_view = "Add A Verif"


'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call check_for_MAXIS(true)
Call MAXIS_case_number_finder(MAXIS_case_number)

interview_date = date & ""
show_err_msg_during_movement = ""
script_run_lowdown = ""
developer_mode = False

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then
		STATS_manualtime = STATS_manualtime + (timer - start_time)
		Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
	End If
End If
If MX_region = "TRAINING" Then developer_mode = True

interview_started_time = time
MFIP_orientation_assessed_and_completed = False

msg_what_script_does_btn = 101
msg_save_your_work_btn = 102
msg_script_interaction_btn = 103
msg_show_instructions_btn = 104
msg_script_messaging_btn = 105
msg_show_quick_start_guide_btn = 106
msg_show_faq_btn = 107
interpreter_servicves_btn = 108

'THIS IS WHERE THE CASE NUMBER DIALOG BELONGS BUT SINCE THIS SCRIPT IS ACCESSED FROM A REDIRECT WE DON'T NEED A REDIRECT

Do
	Call navigate_to_MAXIS_screen("STAT", "SUMM")
	EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"
EMReadScreen case_pw, 7, 21, 17

If CAF_form = "CAF (DHS-5223)" Then CAF_form_name = "Combined Application Form"
If CAF_form = "HUF (DHS-8107)" Then CAF_form_name = "Household Update Form"
If CAF_form = "SNAP App for Srs (DHS-5223F)" Then CAF_form_name = "SNAP Application for Seniors"
If CAF_form = "MNbenefits" Then CAF_form_name = "MNbenefits Web Form"
If CAF_form = "Combined AR for Certain Pops (DHS-3727)" Then CAF_form_name = "Combined Annual Renewal"

If CAF_form = "SNAP App for Srs (DHS-5223F)" OR CAF_form = "Combined AR for Certain Pops (DHS-3727)" Then

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 261, 160, "Unsupported Forms"
	  Text 10, 10, 215, 20, "The following forms have slightly different phrasing from the primary application/recertification forms (like the CAF or HUF):"
	  Text 25, 40, 195, 10, "SNAP App for Srs (DHS-5223F)"
	  Text 25, 50, 190, 10, "Combined AR for Certain Pops (DHS-3727)"
	  Text 10, 70, 240, 20, "We are developing new support for this script as use continues. Until then, the Interview script will function using the phrasing of the CAF form."
	  Text 10, 95, 240, 40, "The requirements of the interview remain the same regardless of the form received. Many questions can be left blank as necessary, the important steps are to document all information received and discussed in the interview. Use the fields in the script to the best of your ability to document the details of the interview."
	  ButtonGroup ButtonPressed
		OkButton 205, 140, 50, 15
	EndDialog

	dialog Dialog1
	Call check_for_MAXIS(False)

End If

If select_err_msg_handling = "Alert at the time you attempt to save each page of the dialog." Then show_err_msg_during_movement = TRUE
If select_err_msg_handling = "Alert only once completing and leaving the final dialog." Then show_err_msg_during_movement = FALSE

show_known_addr = FALSE
vars_filled = FALSE

Call back_to_SELF
Call restore_your_work(vars_filled)			'looking for a 'restart' run

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
EMReadScreen worker_id_for_data_table, 7, 21, 14
EMReadScreen case_name_for_data_table, 25, 21, 40
case_name_for_data_table = trim(case_name_for_data_table)

If snap_status = "APP OPEN" or snap_status = "APP CLOSE" Then snap_status = "ACTIVE"
If grh_status = "APP OPEN" or grh_status = "APP CLOSE" Then grh_status = "ACTIVE"
If mfip_status = "APP OPEN" or mfip_status = "APP CLOSE" Then mfip_status = "ACTIVE"
If dwp_status = "APP OPEN" or dwp_status = "APP CLOSE" Then dwp_status = "ACTIVE"
If ga_status = "APP OPEN" or ga_status = "APP CLOSE" Then ga_status = "ACTIVE"
If msa_status = "APP OPEN" or msa_status = "APP CLOSE" Then msa_status = "ACTIVE"
If vars_filled = False Then
	If adult_cash_case = True Then type_of_cash = "Adult"
	If family_cash_case = True Then type_of_cash = "Family"
	If case_pending = True Then
		Call navigate_to_MAXIS_screen("REPT", "PND2")
		EMReadScreen pnd2_disp_limit, 13, 6, 35
		If pnd2_disp_limit = "Display Limit" Then transmit
		row = 1
		col = 1
		EMSearch MAXIS_case_number, row, col
		If row <> 24 and row <> 0 Then pnd2_row = row
		EMReadScreen CAF_datestamp, 8, pnd2_row, 38
		CAF_datestamp = replace(CAF_datestamp, " ", "/")


		If unknown_cash_pending = True Then CASH_on_CAF_checkbox = checked
		If ga_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If msa_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If mfip_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If dwp_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If snap_status = "PENDING" Then SNAP_on_CAF_checkbox = checked
		If emer_status = "PENDING" Then EMER_on_CAF_checkbox = checked

	End If
	MAXIS_footer_month = CM_mo
	MAXIS_footer_year = CM_yr
	Call navigate_to_MAXIS_screen("STAT", "REVW")
	EMReadScreen next_cash_revw_mo, 2, 9, 37
	EMReadScreen next_cash_revw_yr, 2, 9, 43
	EMReadScreen next_snap_revw_mo, 2, 9, 57
	EMReadScreen next_snap_revw_yr, 2, 9, 63

	If next_cash_revw_mo = "__" Then next_cash_revw_mo = ""
	If next_cash_revw_yr = "__" Then next_cash_revw_yr = ""
	If next_snap_revw_mo = "__" Then next_snap_revw_mo = ""
	If next_snap_revw_yr = "__" Then next_snap_revw_yr = ""

	cash_revw = False
	snap_revw = False

	If next_cash_revw_mo = CM_mo AND next_cash_revw_yr = CM_yr Then cash_revw = True
	If next_cash_revw_mo = CM_plus_1_mo AND next_cash_revw_yr = CM_plus_1_yr Then cash_revw = True
	If next_cash_revw_mo = CM_plus_2_mo AND next_cash_revw_yr = CM_plus_2_yr Then cash_revw = True

	If next_snap_revw_mo = CM_mo AND next_snap_revw_yr = CM_yr Then snap_revw = True
	If next_snap_revw_mo = CM_plus_1_mo AND next_snap_revw_yr = CM_plus_1_yr Then snap_revw = True
	If next_snap_revw_mo = CM_plus_2_mo AND next_snap_revw_yr = CM_plus_2_yr Then snap_revw = True

	If CAF_datestamp = "" Then
		If cash_revw = True Then
			MAXIS_footer_month = next_cash_revw_mo
			MAXIS_footer_year = next_cash_revw_yr
            If next_cash_revw_mo = CM_plus_2_mo AND next_cash_revw_yr = CM_plus_2_yr Then
                MAXIS_footer_month = CM_plus_1_mo
                MAXIS_footer_year = CM_plus_1_yr
            End If
			call back_to_SELF
			Call navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen CAF_datestamp, 8, 13, 37
			CAF_datestamp = replace(CAF_datestamp, " ", "/")
		End If

		If snap_revw = True Then
			MAXIS_footer_month = next_snap_revw_mo
			MAXIS_footer_year = next_snap_revw_yr
            If next_snap_revw_mo = CM_plus_2_mo AND next_snap_revw_yr = CM_plus_2_yr Then
                MAXIS_footer_month = CM_plus_1_mo
                MAXIS_footer_year = CM_plus_1_yr
            End If
			call back_to_SELF
			Call navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen CAF_datestamp, 8, 13, 37
			CAF_datestamp = replace(CAF_datestamp, " ", "/")
		End If
		If CAF_datestamp = "__/__/__" Then CAF_datestamp = ""
	End If
	If cash_revw = True Then CASH_on_CAF_checkbox = checked
	If snap_revw = True Then SNAP_on_CAF_checkbox = checked
End If

BeginDialog Dialog1, 0, 0, 311, 245, "Programs to Interview For"
  EditBox 55, 40, 80, 15, CAF_datestamp
  CheckBox 185, 40, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 225, 40, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 265, 40, 35, 10, "EMER", EMER_on_CAF_checkbox
  EditBox 40, 135, 260, 15, cash_other_req_detail
  EditBox 40, 155, 260, 15, snap_other_req_detail
  EditBox 40, 175, 260, 15, emer_other_req_detail
  ButtonGroup ButtonPressed
    OkButton 200, 225, 50, 15
    CancelButton 255, 225, 50, 15
  Text 10, 10, 265, 10, "We are going to start the interview based on the information listed on the form:"
  Text 20, 25, 155, 10, CAF_form_name
  Text 20, 45, 35, 10, "CAF Date:"
  GroupBox 180, 25, 125, 30, "Programs marked on CAF"
  Text 15, 60, 295, 10, "As a part of the interview, we need to confirm the programs requested (or being reviewed)."
  Text 15, 75, 210, 10, "Confirm with the resident which programs should be assessed:"
  Text 25, 85, 250, 10, "-Update the checkboxes above to reflect what is marked on the CAF Form"
  Text 25, 95, 200, 10, "-Add any verbal request information in the boxes below."
  GroupBox 5, 110, 300, 85, "OTHER Program Requests (not marked on CAF)"
  Text 40, 125, 130, 10, "Explain how the program was requested."
  Text 15, 140, 20, 10, "Cash:"
  Text 15, 160, 20, 10, "SNAP:"
  Text 15, 180, 25, 10, "EMER:"
  Text 10, 200, 295, 25, "We need to know what programs we are assessing in the interview. Take time with the resident to ensure they understand the requests and we complete all information necesssary to complete the interview."
EndDialog

Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		cash_other_req_detail = trim(cash_other_req_detail)
	    snap_other_req_detail = trim(snap_other_req_detail)
	    emer_other_req_detail = trim(emer_other_req_detail)

		program_requested = False
		If CASH_on_CAF_checkbox = checked Then program_requested = True
		If SNAP_on_CAF_checkbox = checked Then program_requested = True
		If EMER_on_CAF_checkbox = checked Then program_requested = True
		If cash_other_req_detail <> "" Then program_requested = True
		If snap_other_req_detail <> "" Then program_requested = True
		If emer_other_req_detail <> "" Then program_requested = True

		If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
		If program_requested = False Then err_msg = err_msg & vbCr & "* We must indicate a program being requested on the form or verbally. Review the request details with the resident."

		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
save_your_work
Call check_for_MAXIS(False)

Call Navigate_to_MAXIS_screen("CASE", "NOTE")               'Now we navigate to CASE:NOTES
too_old_date = DateAdd("D", -1, CAF_datestamp)              'We don't need to read notes from before the CAF date

Call hest_standards(heat_AC_amt, electric_amt, phone_amt, CAF_datestamp)

note_row = 5
Do
	EMReadScreen note_date, 8, note_row, 6                  'reading the note date

	EMReadScreen note_title, 55, note_row, 25               'reading the note header
	note_title = trim(note_title)

	IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
		appt_notc_sent_on = note_date
	ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
		appt_notc_sent_on = note_date
	ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
		EMReadScreen appt_date, 10, note_row, 63
		appt_date = replace(appt_date, "~", "")
		appt_date = trim(appt_date)
		appt_notc_sent_on = note_date
		appt_date_in_note = appt_date
	END IF

	if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

    note_row = note_row + 1
    if note_row = 19 then
        note_row = 5
        PF8
        EMReadScreen check_for_last_page, 9, 24, 14
        If check_for_last_page = "LAST PAGE" Then Exit Do
    End If
    EMReadScreen next_note_date, 8, note_row, 6
    if next_note_date = "        " then Exit Do
Loop until DateDiff("d", too_old_date, next_note_date) <= 0

cash_request = False
snap_request = False
emer_request = False
If CASH_on_CAF_checkbox = checked OR cash_other_req_detail <> "" Then cash_request = True
If SNAP_on_CAF_checkbox = checked OR snap_other_req_detail <> "" Then snap_request = True
If EMER_on_CAF_checkbox = checked OR emer_other_req_detail <> "" Then emer_request = True

If vars_filled = False Then
	If cash_revw = True AND cash_request = True Then the_process_for_cash = "Renewal"
	If snap_revw = True AND snap_request = True Then the_process_for_snap = "Renewal"

	If unknown_cash_pending = True Then the_process_for_cash ="Application"
	If ga_status = "PENDING" Then the_process_for_cash = "Application"
	If msa_status = "PENDING" Then the_process_for_cash = "Application"
	If mfip_status = "PENDING" Then the_process_for_cash = "Application"
	If dwp_status = "PENDING" Then the_process_for_cash = "Application"
	If snap_status = "PENDING" Then the_process_for_snap = "Application"
	the_process_for_emer = "Application"
End If

dlg_len = 50
y_pos = 25
If cash_request = True Then dlg_len = dlg_len + 20
If snap_request = True Then dlg_len = dlg_len + 20
If emer_request = True Then dlg_len = dlg_len + 20
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 205, dlg_len, "CAF Process"
  Text 10, 10, 35, 10, "Program"
  Text 80, 10, 50, 10, "CAF Process"
  Text 155, 10, 50, 10, "Recert MM/YY"
  If cash_request = True Then
	  Text 10, y_pos + 5, 20, 10, "Cash"
	  DropListBox 35, y_pos, 35, 45, "?"+chr(9)+"Family"+chr(9)+"Adult", type_of_cash
	  DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_cash
	  EditBox 155, y_pos, 20, 15, next_cash_revw_mo
	  EditBox 180, y_pos, 20, 15, next_cash_revw_yr
	  y_pos = y_pos + 20
  End If
  If snap_request = True Then
	  Text 10, y_pos + 5, 20, 10, "SNAP"
	  DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_snap
	  EditBox 155, y_pos, 20, 15, next_snap_revw_mo
	  EditBox 180, y_pos, 20, 15, next_snap_revw_yr
	  y_pos = y_pos + 20
  End If
  If emer_request = True Then
	  Text 10, y_pos + 5, 20, 10, "EMER"
	  DropListBox 35, y_pos, 35, 45, "?"+chr(9)+"EA"+chr(9)+"EGA", type_of_emer
	  DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application", the_process_for_emer
	  y_pos = y_pos + 20
  End If
  y_pos = y_pos + 5
  ButtonGroup ButtonPressed
	OkButton 150, y_pos, 50, 15
EndDialog

Do
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		If len(next_cash_revw_yr) = 4 AND left(next_cash_revw_yr, 2) = "20" Then next_cash_revw_yr = right(next_cash_revw_yr, 2)
		If len(next_snap_revw_yr) = 4 AND left(next_snap_revw_yr, 2) = "20" Then next_snap_revw_yr = right(next_snap_revw_yr, 2)
		If cash_request = True Then
			If the_process_for_cash = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the CASH program is at application or renewal."
			If the_process_for_cash = "Renewal" AND (len(next_cash_revw_mo) <> 2 or len(next_cash_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For CASH at renewal, enter the footer month and year the of the renewal."
		End If
		If snap_request = True Then
			If the_process_for_snap = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the SNAP program is at application or renewal."
			If the_process_for_snap = "Renewal" AND (len(next_snap_revw_mo) <> 2 or len(next_snap_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For SNAP at renewal, enter the footer month and year the of the renewal."
		End If
		If emer_request = True Then
			If type_of_emer = "?" Then r_msg = err_msg & vbNewLine & "*Indicate if EMER request in EA or EGA"
		End If


		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
Call check_for_MAXIS(False)

If the_process_for_snap = "Application" Then expedited_determination_needed = True
If snap_status = "PENDING" Then expedited_determination_needed = True
If type_of_cash = "Adult" Then family_cash_case_yn = "No"
If type_of_cash = "Family" Then family_cash_case_yn = "Yes"
If vars_filled = TRUE Then show_known_addr = TRUE		'This is a setting for the address dialog to see the view

Call convert_date_into_MAXIS_footer_month(CAF_datestamp, MAXIS_footer_month, MAXIS_footer_year)
original_footer_month = MAXIS_footer_month
original_footer_year = MAXIS_footer_year

'If we already know the variables because we used 'restore your work' OR if there is no case number, we don't need to read the information from MAXIS
If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then
	Call back_to_SELF

	Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)			'Then we create an array of the the full hh list for looping purpoases
	full_hh_list = Split(list_for_array, chr(9))


	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
    EMWriteScreen "01", 20, 76
    transmit
    EMReadScreen id_ver_code, 2, 9, 68
	If id_ver_code <> "__" AND id_ver_code <> "NO" Then applicant_id_on_file_yn = "Yes"
	If id_ver_code = "__" OR id_ver_code = "NO" Then applicant_id_on_file_yn = "No"
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadscreen ref_nbr, 2, 4, 33
		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			EMWaitReady 0, 0
		End If
		If client_array <> "" Then client_array = client_array & "|" & ref_nbr
		If client_array = "" Then client_array = client_array & ref_nbr
		transmit      'Going to the next MEMB panel
		Emreadscreen edit_check, 7, 24, 2 'looking to see if we are at the last member
		member_count = member_count + 1
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
	' MsgBox client_array
	client_array = split(client_array, "|")

	clt_count = 0

	For each hh_clt in client_array

		ReDim Preserve HH_MEMB_ARRAY(last_const, clt_count)
		HH_MEMB_ARRAY(ref_number, clt_count) = hh_clt
        HH_MEMB_ARRAY(pers_in_maxis, clt_count) = True
        HH_MEMB_ARRAY(ignore_person, clt_count) = False
		' HH_MEMB_ARRAY(define_the_member, clt_count)

		Call navigate_to_MAXIS_screen("STAT", "MEMB")		'===============================================================================================
		EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			EMWaitReady 0, 0
			HH_MEMB_ARRAY(last_name_const, clt_count) = "UNABLE TO FIND"
			HH_MEMB_ARRAY(first_name_const, clt_count) = "Access Denied"
			HH_MEMB_ARRAY(mid_initial, clt_count) = ""
			HH_MEMB_ARRAY(access_denied, clt_count) = TRUE
		Else
			HH_MEMB_ARRAY(access_denied, clt_count) = FALSE
			EMReadscreen HH_MEMB_ARRAY(last_name_const, clt_count), 25, 6, 30
			EMReadscreen HH_MEMB_ARRAY(first_name_const, clt_count), 12, 6, 63
			EMReadscreen HH_MEMB_ARRAY(mid_initial, clt_count), 1, 6, 79
			EMReadScreen HH_MEMB_ARRAY(age, clt_count), 3, 8, 76

			EMReadScreen HH_MEMB_ARRAY(date_of_birth, clt_count), 10, 8, 42
			EMReadScreen HH_MEMB_ARRAY(ssn, clt_count), 11, 7, 42
			EMReadScreen HH_MEMB_ARRAY(ssn_verif, clt_count), 1, 7, 68
			EMReadScreen HH_MEMB_ARRAY(birthdate_verif, clt_count), 2, 8, 68
			EMReadScreen HH_MEMB_ARRAY(gender, clt_count), 1, 9, 42
			EMReadScreen HH_MEMB_ARRAY(race, clt_count), 30, 17, 42
			EMReadScreen HH_MEMB_ARRAY(spoken_lang, clt_count), 20, 12, 42
			EMReadScreen HH_MEMB_ARRAY(written_lang, clt_count), 29, 13, 42
			EMReadScreen HH_MEMB_ARRAY(interpreter, clt_count), 1, 14, 68
			EMReadScreen HH_MEMB_ARRAY(alias_yn, clt_count), 1, 15, 42
			EMReadScreen HH_MEMB_ARRAY(ethnicity_yn, clt_count), 1, 16, 68

			HH_MEMB_ARRAY(age, clt_count) = trim(HH_MEMB_ARRAY(age, clt_count))
			If HH_MEMB_ARRAY(age, clt_count) = "" Then HH_MEMB_ARRAY(age, clt_count) = 0
			HH_MEMB_ARRAY(age, clt_count) = HH_MEMB_ARRAY(age, clt_count) * 1

			HH_MEMB_ARRAY(last_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(last_name_const, clt_count), "_", ""))
			HH_MEMB_ARRAY(first_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(first_name_const, clt_count), "_", ""))
			HH_MEMB_ARRAY(mid_initial, clt_count) = replace(HH_MEMB_ARRAY(mid_initial, clt_count), "_", "")
			HH_MEMB_ARRAY(full_name_const, clt_count) = HH_MEMB_ARRAY(first_name_const, clt_count) & " " & HH_MEMB_ARRAY(last_name_const, clt_count)
			EMReadScreen HH_MEMB_ARRAY(id_verif, clt_count), 2, 9, 68

			EMReadScreen HH_MEMB_ARRAY(rel_to_applcnt, clt_count), 2, 10, 42              'reading the relationship from MEMB'
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "01" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Self"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "02" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Spouse"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "03" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "04" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "05" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "06" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "08" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "09" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "10" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Aunt"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "11" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Uncle"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "12" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Niece"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "13" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Nephew"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "14" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Cousin"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "15" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Grandparent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "16" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Grandchild"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "17" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Other Relative"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "18" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Legal Guardian"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "24" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Not Related"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "25" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Live-in Attendant"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "27" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Unknown"

			If HH_MEMB_ARRAY(id_verif, clt_count) = "BC" Then HH_MEMB_ARRAY(id_verif, clt_count) = "BC - Birth Certificate"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "RE" Then HH_MEMB_ARRAY(id_verif, clt_count) = "RE - Religious Record"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DL - Drivers License/ST ID"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DV - Divorce Decree"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AL - Alien Card"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AD" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AD - Arrival//Depart"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DR" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DR - Doctor Stmt"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "PV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "PV - Passport/Visa"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "OT" Then HH_MEMB_ARRAY(id_verif, clt_count) = "OT - Other Document"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "NO" Then HH_MEMB_ARRAY(id_verif, clt_count) = "NO - No Veer Prvd"

			If HH_MEMB_ARRAY(age, clt_count) > 18 then
				HH_MEMB_ARRAY(cash_minor, clt_count) = FALSE
			Else
				HH_MEMB_ARRAY(cash_minor, clt_count) = TRUE
			End If
			If HH_MEMB_ARRAY(age, clt_count) > 21 then
				HH_MEMB_ARRAY(snap_minor, clt_count) = FALSE
			Else
				HH_MEMB_ARRAY(snap_minor, clt_count) = TRUE
			End If

			HH_MEMB_ARRAY(date_of_birth, clt_count) = replace(HH_MEMB_ARRAY(date_of_birth, clt_count), " ", "/")
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "BC" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "BC - Birth Certificate"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "RE" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "RE - Religious Record"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DL" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DL - Drivers License/State ID"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DV" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DV - Divorce Decree"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "AL" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "AL - Alien Card"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DR" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DR - Doctor Statement"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "OT" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "OT - Other Document"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "PV" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "PV - Passport/Visa"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "NO" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "NO - No Verif Provided"

			HH_MEMB_ARRAY(ssn, clt_count) = replace(HH_MEMB_ARRAY(ssn, clt_count), " ", "-")
			if HH_MEMB_ARRAY(ssn, clt_count) = "___-__-____" Then HH_MEMB_ARRAY(ssn, clt_count) = ""
			HH_MEMB_ARRAY(ssn_no_space, clt_count) = replace(HH_MEMB_ARRAY(ssn, clt_count), "-", "")

			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "A" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "A - SSN Applied For"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "P" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "P - SSN Provided, verif Pending"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "N" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "N - SSN Not Provided"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "V" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "V - SSN Verified via Interface"

			If HH_MEMB_ARRAY(gender, clt_count) = "M" Then HH_MEMB_ARRAY(gender, clt_count) = "Male"
			If HH_MEMB_ARRAY(gender, clt_count) = "F" Then HH_MEMB_ARRAY(gender, clt_count) = "Female"

			HH_MEMB_ARRAY(race, clt_count) = trim(HH_MEMB_ARRAY(race, clt_count))

			HH_MEMB_ARRAY(spoken_lang, clt_count) = replace(replace(HH_MEMB_ARRAY(spoken_lang, clt_count), "_", ""), "  ", " - ")
			HH_MEMB_ARRAY(written_lang, clt_count) = trim(replace(replace(replace(HH_MEMB_ARRAY(written_lang, clt_count), "_", ""), "  ", " - "), "(HRF)", ""))


			Call navigate_to_MAXIS_screen("STAT", "MEMI")		'===============================================================================================
			EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
			transmit

			EMReadScreen HH_MEMB_ARRAY(marital_status, clt_count), 1, 7, 40
			EMReadScreen HH_MEMB_ARRAY(spouse_ref, clt_count), 2, 9, 49
			EMReadScreen HH_MEMB_ARRAY(spouse_name, clt_count), 40, 9, 52
			EMReadScreen HH_MEMB_ARRAY(last_grade_completed, clt_count), 2, 10, 49
			EMReadScreen HH_MEMB_ARRAY(citizen, clt_count), 1, 11, 49
			EMReadScreen HH_MEMB_ARRAY(other_st_FS_end_date, clt_count), 8, 13, 49
			EMReadScreen HH_MEMB_ARRAY(in_mn_12_mo, clt_count), 1, 14, 49
			EMReadScreen HH_MEMB_ARRAY(residence_verif, clt_count), 1, 14, 78
			EMReadScreen HH_MEMB_ARRAY(mn_entry_date, clt_count), 8, 15, 49
			EMReadScreen HH_MEMB_ARRAY(former_state, clt_count), 2, 15, 78

			If HH_MEMB_ARRAY(marital_status, clt_count) = "N" Then HH_MEMB_ARRAY(marital_status, clt_count) = "N - Never Married"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "M" Then HH_MEMB_ARRAY(marital_status, clt_count) = "M - Married Living with Spouse"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "S" Then HH_MEMB_ARRAY(marital_status, clt_count) = "S - Married Living Apart"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "L" Then HH_MEMB_ARRAY(marital_status, clt_count) = "L - Legally Seperated"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "D" Then HH_MEMB_ARRAY(marital_status, clt_count) = "D - Divorced"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "W" Then HH_MEMB_ARRAY(marital_status, clt_count) = "W - Widowed"
			If HH_MEMB_ARRAY(spouse_ref, clt_count) = "__" Then HH_MEMB_ARRAY(spouse_ref, clt_count) = ""
			HH_MEMB_ARRAY(spouse_name, clt_count) = trim(HH_MEMB_ARRAY(spouse_name, clt_count))

			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "00" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Not Attended or Pre-Grade 1 - 00"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "12" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "High School Diploma or GED - 12"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "13" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Some Post Sec Education - 13"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "14" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "High School Plus Certiificate - 14"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "15" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Four Year Degree - 15"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "16" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Grad Degree - 16"
			If len(HH_MEMB_ARRAY(last_grade_completed, clt_count)) = 2 Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Grade " & HH_MEMB_ARRAY(last_grade_completed, clt_count)
			If HH_MEMB_ARRAY(citizen, clt_count) = "Y" Then HH_MEMB_ARRAY(citizen, clt_count) = "Yes"
			If HH_MEMB_ARRAY(citizen, clt_count) = "N" Then HH_MEMB_ARRAY(citizen, clt_count) = "No"

			HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = replace(HH_MEMB_ARRAY(other_st_FS_end_date, clt_count), " ", "/")
			If HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = "__/__/__" Then HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = ""
			If HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "Y" Then HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "Yes"
			If HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "N" Then HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "No"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "1" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "1 - Rent Receipt"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "2" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "2 - Landlord's Statement"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "3" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "3 - Utility Bill"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "4" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "4 - Other"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "N" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "N - Verif Not Provided"
			HH_MEMB_ARRAY(mn_entry_date, clt_count) = replace(HH_MEMB_ARRAY(mn_entry_date, clt_count), " ", "/")
			If HH_MEMB_ARRAY(mn_entry_date, clt_count) = "__/__/__" Then HH_MEMB_ARRAY(mn_entry_date, clt_count) = ""
			If HH_MEMB_ARRAY(former_state, clt_count) = "__" Then HH_MEMB_ARRAY(former_state, clt_count) = ""


		End If

		memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)
		If HH_MEMB_ARRAY(fs_pwe, clt_count) = "Yes" Then the_pwe_for_this_case = HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)

		' HH_MEMB_ARRAY(clt_count).intend_to_reside_in_mn = "Yes"

		' ReDim Preserve ALL_ANSWERS_ARRAY(ans_notes, clt_count)
		clt_count = clt_count + 1
	Next

	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		HH_MEMB_ARRAY(race_a_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_b_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_n_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_p_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_w_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(snap_req_checkbox, the_members) = unchecked
		If SNAP_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(snap_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(cash_req_checkbox, the_members) = unchecked
		If CASH_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(emer_req_checkbox, the_members) = unchecked
		If EMER_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(emer_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked

		HH_MEMB_ARRAY(clt_has_sponsor, the_members) = ""
		HH_MEMB_ARRAY(client_verification, the_members) = ""
		HH_MEMB_ARRAY(client_verification_details, the_members) = ""
		HH_MEMB_ARRAY(client_notes, the_members) = ""
		HH_MEMB_ARRAY(imig_status, the_members) = ""
	Next

	'Now we gather the address information that exists in MAXIS
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, reservation_name, mail_line_one, mail_line_two, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, addr_eff_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_one_type, phone_two_type, phone_three_type, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

	arep_in_MAXIS = False
	arep_exists = False
	update_arep = True
	Call access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)
	If arep_name <> "" Then
		If arep_phone_two <> "" Then arep_phone_number = arep_phone_two
		If arep_phone_one <> "" Then arep_phone_number = arep_phone_one
		MAXIS_arep_name = arep_name
		MAXIS_arep_relationship = arep_relationship
		MAXIS_arep_phone_number = arep_phone_number
		MAXIS_arep_addr_street = arep_addr_street
		MAXIS_arep_addr_city = arep_addr_city
		MAXIS_arep_addr_state = arep_addr_state
		MAXIS_arep_addr_zip = arep_addr_zip
		arep_in_MAXIS = True
		MAXIS_arep_updated = False
		arep_exists = True
		update_arep = False
		MAXIS_arep_complete_forms_checkbox = checked
	End If
	If forms_to_arep = "Y" Then arep_get_notices_checkbox = checked

	show_known_addr = True
End If

'Giving the buttons specific unumerations so they don't think they are eachother
next_btn					= 100
' back_btn					= 1010
update_information_btn		= 1020
save_information_btn		= 1030
clear_mail_addr_btn			= 1040
clear_phone_one_btn			= 1041
clear_phone_two_btn			= 1042
clear_phone_three_btn		= 1043
add_person_btn				= 1050

caf_page_one_btn			= 1300
caf_addr_btn				= 1400
caf_membs_btn				= 1500
case_details_btn			= 1600
caf_qual_q_btn				= 2200
caf_last_page_btn			= 2300
finish_interview_btn		= 2400
exp_income_guidance_btn 	= 2500
open_hsr_manual_transfer_page_btn = 2610
verif_button				= 2800
expedited_determination_btn	= 3010
return_btn 					= 900
enter_btn					= 901
continue_btn				= 902
done_btn					= 903
review_btn					= 904
finish_btn					= 905
clear_btn					= 906
fill_button					= 907
calculate_btn				= 908
update_btn					= 909
add_verif_button			= 910

add_another_jobs_btn			= 800
remove_one_jobs_btn				= 801
add_another_busi_btn			= 802
remove_one_busi_btn				= 803
add_another_unea_btn			= 804
remove_one_unea_btn				= 805
add_another_btn					= 806
remove_one_btn					= 807
income_calc_btn					= 808
asset_calc_btn					= 809
housing_calc_btn				= 810
utility_calc_btn				= 811
ht_id_in_solq_btn				= 812
snap_active_in_another_state_btn	= 813
case_previously_had_postponed_verifs_btn = 814
household_in_a_facility_btn		= 815
knowledge_now_support_btn		= 816
te_02_10_01_btn					= 817
cm_04_12_btn					= 818
ebt_card_info_btn				= 819
hsr_manual_expedited_snap_btn	= 820
hsr_applications_btn		= 821
sir_exp_flowchart_btn			= 822
ryb_exp_identity_btn			= 823
ryb_exp_timeliness_btn			= 824
cm_04_04_btn					= 825
cm_04_06_btn					= 826
amounts_btn						= 827
determination_btn				= 828
return_to_dialog_button			= 829
fn_review_btn					= 830

btn_placeholder = 4000
For btn_count = 0 to UBound(HH_MEMB_ARRAY, 2)
	HH_MEMB_ARRAY(button_one, btn_count) = 500 + btn_count
	HH_MEMB_ARRAY(button_two, btn_count) = 600 + btn_count

	If HH_MEMB_ARRAY(age, btn_count) < 18 Then children_under_18_in_hh = True
	If HH_MEMB_ARRAY(age, btn_count) < 22 Then children_under_22_in_hh = True
	If HH_MEMB_ARRAY(age, btn_count) > 4 AND HH_MEMB_ARRAY(age, btn_count) < 18 Then school_age_children_in_hh = True
Next
interview_date = interview_date & ""
selected_memb = 0
pick_a_client = replace(all_the_clients, "Select or Type", "Select One...")

interview_questions_clear = False
Do
	Do
		Do
			Do

			Dialog1 = Empty
				call define_main_dialog

				err_msg = ""

				prev_page = page_display
				previous_button_pressed = ButtonPressed

				dialog Dialog1
				save_your_work
				cancel_confirmation
				Call assess_caf_1_expedited_questions(expedited_screening)
				Call verification_dialog
				Call check_for_errors(interview_questions_clear)
                If ButtonPressed = interpreter_servicves_btn Then
                    run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://itwebpw026/content/forms/af/_internal/hhs/human_services/initial_contact_access/AF10196.html"
                Else
                    Call display_errors(err_msg, False, show_err_msg_during_movement)
                End If

				If snap_status <> "ACTIVE" Then Call evaluate_for_expedited(intv_app_month_income, intv_app_month_asset, intv_app_month_housing_expense, intv_exp_pay_heat_checkbox, intv_exp_pay_ac_checkbox, intv_exp_pay_electricity_checkbox, intv_exp_pay_phone_checkbox, app_month_utilities_cost, app_month_expenses, case_is_expedited)
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

Call check_for_MAXIS(False)

If relative_caregiver_yn = "Yes" Then absent_parent_yn = "Yes"
exp_pregnant_who = trim(exp_pregnant_who)
If exp_pregnant_who = "Select or Type" Then exp_pregnant_who = ""

for each_member = 0 to UBound(HH_MEMB_ARRAY, 2)
	If HH_MEMB_ARRAY(id_verif, each_member) = "Found in SOLQ/SMI" Then HH_MEMB_ARRAY(id_verif, each_member) = "Identity verified per Verify MN interface"
next

'Navigate back to self and to EDRS
Back_to_self
CALL navigate_to_MAXIS_screen("INFC", "EDRS")
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN
	STATS_manualtime = STATS_manualtime + (timer - start_time + add_to_time)
	script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
End If

edrs_match_found = False
For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
    If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
    	'Write in SSN number into EDRS
    	EMwritescreen HH_MEMB_ARRAY(ssn_no_space, the_memb), 2, 7
    	transmit
    	Emreadscreen SSN_output, 7, 24, 2

    	'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
    	IF SSN_output = "NO DISQ" THEN
    		EMWritescreen HH_MEMB_ARRAY(last_name_const, the_memb), 2, 24
    		EMWritescreen HH_MEMB_ARRAY(first_name_const, the_memb), 2, 58
    		EMWritescreen HH_MEMB_ARRAY(mid_initial, the_memb), 2, 76
    		transmit
    		EMreadscreen NAME_output, 7, 24, 2
    		IF NAME_output = "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name.
    			HH_MEMB_ARRAY(edrs_msg, the_memb) = "No disqualifications found for Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb)
    			HH_MEMB_ARRAY(edrs_match, the_memb) = FALSE
    		ELSE
    			HH_MEMB_ARRAY(edrs_msg, the_memb) = "Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb) & " has a potential name match."
    			HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE
    			edrs_match_found = True
    		END IF
    	ELSE
    		HH_MEMB_ARRAY(edrs_msg, the_memb) = "Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb) & " has SSN Match."    'If after searching a SSN number you don't get the NO DISQ message then let worker know you found the SSN
    		HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE
    		edrs_match_found = True
    	END IF
		STATS_manualtime = STATS_manualtime + 49
    End If
Next

Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "EDRs Search Review"
		  ButtonGroup ButtonPressed
		    PushButton 465, 360, 80, 15, "Continue", continue_btn
		    Text 10, 10, 320, 10, "EDRs has been completed for all Household Members."
			y_pos = 25
		    For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
				If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
                    Text 20, y_pos, 420, 10, HH_MEMB_ARRAY(edrs_msg, the_memb)

    				PushButton 390, y_pos, 70, 10, "SSN SEARCH", HH_MEMB_ARRAY(button_one, the_memb)
    				PushButton 460, y_pos, 70, 10, "NAME SEARCH", HH_MEMB_ARRAY(button_two, the_memb)
    				If HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE Then
    					Text 30, y_pos + 20, 45, 10, "EDRs Notes:"
    		  		    EditBox 80, y_pos + 15, 450, 15, HH_MEMB_ARRAY(edrs_notes, the_memb)
    					y_pos = y_pos + 20
    				End If
    				y_pos = y_pos + 20
                End If
			Next
		    Text 15, 350, 70, 10, "EDRs CASE Notes:"
		    EditBox 15, 360, 440, 15, edrs_notes_for_case
		EndDialog

		dialog Dialog1

		cancel_confirmation
		For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
			If ButtonPressed = HH_MEMB_ARRAY(button_one, the_memb) OR ButtonPressed = HH_MEMB_ARRAY(button_two, the_memb) Then
				err_msg = err_msg & "LOOP"
				EMReadScreen edrs_check, 12, 1, 36
				If edrs_check <> "EDRS Inquiry" Then
					Back_to_self
					CALL navigate_to_MAXIS_screen("INFC", "EDRS")
				End If
				If ButtonPressed = HH_MEMB_ARRAY(button_two, the_memb) Then
					EMWritescreen HH_MEMB_ARRAY(last_name_const, the_memb), 2, 24
					EMWritescreen HH_MEMB_ARRAY(first_name_const, the_memb), 2, 58
					EMWritescreen HH_MEMB_ARRAY(mid_initial, the_memb), 2, 76
				End If
				If ButtonPressed = HH_MEMB_ARRAY(button_one, the_memb) Then EMwritescreen HH_MEMB_ARRAY(ssn_no_space, the_memb), 2, 7
				transmit
			End If
		Next

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

call back_to_SELF

Call check_for_MAXIS(False)

CAF_MONTH_DATE = MAXIS_footer_month & "/1/" & MAXIS_footer_year
CAF_MONTH_DATE = DateAdd("d", 0, CAF_MONTH_DATE)
MONTH_BEFORE_CAF = DateAdd("m", -1, CAF_MONTH_DATE)
MONTH_AFTER_CAF = DateAdd("m", 1, CAF_MONTH_DATE)

APPLICATION_MONTH = MAXIS_footer_month
APPLICATION_YEAR = MAXIS_footer_year
CASH_NEXT_REVW_MONTH = ""
CASH_NEXT_REVW_YEAR = ""
CASH_REVW_DATE = ""
SNAP_NEXT_REVW_MONTH = ""
SNAP_NEXT_REVW_YEAR = ""
SNAP_REVW_DATE = ""
cash_revw_due = False
snap_revw_due = False

revw_panel_interview_date = ""
If case_active = True Then
	Call navigate_to_MAXIS_screen("STAT", "REVW")
	If ga_status = "ACTIVE" OR msa_status = "ACTIVE" OR mfip_status = "ACTIVE" OR grh_status = "ACTIVE" Then
		EMReadScreen CASH_NEXT_REVW_MONTH, 2, 9, 37
		EMReadScreen CASH_NEXT_REVW_YEAR, 2, 9, 43
		CASH_REVW_DATE = CASH_NEXT_REVW_MONTH & "/1/" & CASH_NEXT_REVW_YEAR
		CASH_REVW_DATE = DateAdd("d", 0, CASH_REVW_DATE)
		If DateDiff("d", CASH_REVW_DATE, CAF_MONTH_DATE) = 0 Then cash_revw_due = True
		If DateDiff("d", CASH_REVW_DATE, MONTH_AFTER_CAF) = 0 Then cash_revw_due = True
	End If
	If snap_status = "ACTIVE" Then
		EMReadScreen SNAP_NEXT_REVW_MONTH, 2, 9, 57
		EMReadScreen SNAP_NEXT_REVW_YEAR, 2, 9, 63
		SNAP_REVW_DATE = SNAP_NEXT_REVW_MONTH & "/1/" & SNAP_NEXT_REVW_YEAR
		SNAP_REVW_DATE = DateAdd("d", 0, SNAP_REVW_DATE)
		If DateDiff("d", SNAP_REVW_DATE, CAF_MONTH_DATE) = 0 Then snap_revw_due = True
		If DateDiff("d", SNAP_REVW_DATE, MONTH_AFTER_CAF) = 0 Then snap_revw_due = True
	End If

End If

prog_cash_1_intvw_date = ""
prog_cash_2_intvw_date = ""
prog_emer_intvw_date = ""
prog_grh_intvw_date = ""
prog_snap_intvw_date = ""
update_prog = False
If case_pending = True Then
	Call navigate_to_MAXIS_screen("STAT", "PROG")

	EMReadScreen prog_cash_1_status, 4, 6, 74
	If prog_cash_1_status = "PEND" Then
		EMReadScreen prog_cash_1_intvw_date, 8, 6, 55
		prog_cash_1_intvw_date = replace(prog_cash_1_intvw_date, " ", "/")
		If prog_cash_1_intvw_date = "__/__/__" Then prog_cash_1_intvw_date = ""
		If prog_cash_1_intvw_date = "" Then update_prog = True
	End If
	EMReadScreen prog_cash_2_status, 4, 7, 74
	If prog_cash_2_status = "PEND" Then
		EMReadScreen prog_cash_2_intvw_date, 8, 7, 55
		prog_cash_2_intvw_date = replace(prog_cash_2_intvw_date, " ", "/")
		If prog_cash_2_intvw_date = "__/__/__" Then prog_cash_2_intvw_date = ""
		If prog_cash_2_intvw_date = "" Then update_prog = True
	End If
	EMReadScreen prog_emer_status, 4, 8, 74
	If prog_emer_status = "PEND" Then
		EMReadScreen prog_emer_intvw_date, 8, 8, 55
		prog_emer_intvw_date = replace(prog_emer_intvw_date, " ", "/")
		If prog_emer_intvw_date = "__/__/__" Then prog_emer_intvw_date = ""
		If prog_emer_intvw_date = "" Then update_prog = True
	End If
	EMReadScreen prog_grh_status, 4, 9, 74
	If prog_grh_status = "PEND" Then
		EMReadScreen prog_grh_intvw_date, 8, 9, 55
		prog_grh_intvw_date = replace(prog_grh_intvw_date, " ", "/")
		If prog_grh_intvw_date = "__/__/__" Then prog_grh_intvw_date = ""
		If prog_grh_intvw_date = "" Then update_prog = True
	End If
	EMReadScreen prog_snap_status, 4, 10, 74
	If prog_snap_status = "PEND" Then
		EMReadScreen prog_snap_intvw_date, 8, 10, 55
		prog_snap_intvw_date = replace(prog_snap_intvw_date, " ", "/")
		If prog_snap_intvw_date = "__/__/__" Then prog_snap_intvw_date = ""
		If prog_snap_intvw_date = "" Then update_prog = True
	End If
End If

update_revw = False
If cash_revw_due = True OR snap_revw_due = True Then
	Call back_to_SELF
	If cash_revw_due = True Then
		MAXIS_footer_month = CASH_NEXT_REVW_MONTH
		MAXIS_footer_year = CASH_NEXT_REVW_YEAR
		Call navigate_to_MAXIS_screen("STAT", "REVW")

		EMReadScreen cash_revw_status_code, 1, 7, 40
		If cash_revw_status_code = "N" OR cash_revw_status_code = "I" OR cash_revw_status_code = "U" Then
			EMReadScreen revw_panel_interview_date, 8, 15, 37
			revw_panel_interview_date = replace(revw_panel_interview_date, " ", "/")
			If revw_panel_interview_date = "__/__/__" Then revw_panel_interview_date = ""
			If revw_panel_interview_date = "" Then update_revw = True
		End If

		MAXIS_footer_month = original_footer_month
		MAXIS_footer_year = original_footer_year
	End If

	Call back_to_SELF
	If snap_revw_due = True Then
		MAXIS_footer_month = SNAP_NEXT_REVW_MONTH
		MAXIS_footer_year = SNAP_NEXT_REVW_YEAR
		Call navigate_to_MAXIS_screen("STAT", "REVW")

		EMReadScreen snap_revw_status_code, 1, 7, 60
		If snap_revw_status_code = "N" OR snap_revw_status_code = "I" OR snap_revw_status_code = "U" Then
			EMReadScreen revw_panel_interview_date, 8, 15, 37
			revw_panel_interview_date = replace(revw_panel_interview_date, " ", "/")
			If revw_panel_interview_date = "__/__/__" Then revw_panel_interview_date = ""
			If revw_panel_interview_date = "" Then update_revw = True
		End If

		MAXIS_footer_month = original_footer_month
		MAXIS_footer_year = original_footer_year
	End If
End If
Call back_to_SELF


If update_revw = True OR update_prog = True Then
	If update_revw = True OR update_prog = True Then dlg_len = 300
	If update_revw = False OR update_prog = True Then dlg_len = 170
	If update_revw = True OR update_prog = False Then dlg_len = 190
	y_pos = 40
	confirm_update_revw = 0
	confirm_update_prog = 0

	If update_revw = True Then confirm_update_revw = 1
	If update_prog = True Then confirm_update_prog = 1
	If prog_cash_1_status = "PEND" AND prog_cash_1_intvw_date = "" Then prog_update_cash_1_checkbox = checked
	If prog_cash_2_status = "PEND" AND prog_cash_2_intvw_date = "" Then prog_update_cash_2_checkbox = checked
	If prog_emer_status = "PEND" AND prog_emer_intvw_date = "" Then prog_update_emer_checkbox = checked
	If prog_grh_status = "PEND" AND prog_grh_intvw_date = "" Then prog_update_grh_checkbox = checked
	If prog_snap_status = "PEND" AND prog_snap_intvw_date = "" Then prog_update_snap_checkbox = checked

	BeginDialog Dialog1, 0, 0, 251, dlg_len, "Update Interview Date in STAT"
	  Text 10, 10, 235, 25, "It appears that the interview date needs to be added to STAT panels. Since the interview is now completed, the script can upate the correct panels with the interview date."
	  If update_revw = True Then
		  GroupBox 5, y_pos, 240, 125, "STAT/REVW Needs to be Updated"
		  OptionGroup RadioGroupREVW
		    RadioButton 10, y_pos + 15, 185, 10, "YES! Update REVW with the Interview Date/CAF Date", confirm_update_revw
		    RadioButton 10, y_pos + 80, 100, 10, "No, do not update REVW", do_not_update_revw
		  Text 20, y_pos + 30, 125, 10, "Interview Date: " & interview_date
		  Text 35, y_pos + 40, 95, 10, "CAF Date: " & CAF_datestamp
		  Text 20, y_pos + 55, 175, 20, "If the REVW Status has not been updated already, it will be changed to an 'I' when the dates are entered."
		  Text 20, y_pos + 95, 220, 10, "Reason REVW should not be updated with the Interview/CAF Date:"
		  EditBox 20, y_pos + 105, 220, 15, no_update_revw_reason
		  y_pos = 170
	  End If
	  If update_prog = True Then
		  GroupBox 5, y_pos, 240, 105, "STAT/PROG Needs to be Updated"
		  OptionGroup RadioGroupPROG
		    RadioButton 10, y_pos + 15, 200, 10, "YES! Update PROG with the Interview Date " & interview_date, confirm_update_prog
		    RadioButton 10, y_pos + 60, 90, 10, "No, do not update PROG", do_not_update_prog
		  CheckBox 25, y_pos + 25, 40, 10, "CASH 1", prog_update_cash_1_checkbox
		  CheckBox 25, y_pos + 35, 40, 10, "CASH 2", prog_update_cash_2_checkbox
		  CheckBox 25, y_pos + 45, 30, 10, "EMER", prog_update_emer_checkbox
		  CheckBox 85, y_pos + 25, 30, 10, "GRH", prog_update_grh_checkbox
		  CheckBox 85, y_pos + 35, 30, 10, "SNAP", prog_update_snap_checkbox
		  Text 20, y_pos + 75, 200, 10, "Reason PROG should not be updated with the Interview Date:"
		  EditBox 20, y_pos + 85, 220, 15, no_update_prog_reason
	  End If
	  ButtonGroup ButtonPressed
	    OkButton 195, dlg_len - 20, 50, 15
	EndDialog

	'Running the dialog
	Do
		Do
			err_msg = ""
			Dialog Dialog1
			If update_revw = True Then
				'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
				If do_not_update_revw = 1 AND no_update_revw_reason = "" Then err_msg = err_msg & vbNewLine & "* If REVW is not to be updated, please explain why REVW should not be updated."
			End If

			If update_prog = True Then
				'Requiring a reason for not updating PROG and making sure if confirm is updated that a program is selected.
				If do_not_update_prog = 1 AND no_update_prog_reason = "" Then err_msg = err_msg & vbNewLine & "* If PROG is not to be updated, please explain why PROG should not be updated."
				IF confirm_update_prog = 1 Then
					If prog_update_cash_1_checkbox = unchecked AND prog_update_cash_2_checkbox = unchecked AND prog_update_emer_checkbox = unchecked AND prog_update_grh_checkbox = unchecked AND prog_update_snap_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Select which program to be updated on PROG."
				End If
			End If

			If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
	Call check_for_MAXIS(False)

	intv_mo = DatePart("m", interview_date)     'Setting the date parts to individual variables for ease of writing
	intv_day = DatePart("d", interview_date)
	intv_yr = DatePart("yyyy", interview_date)

	intv_mo = right("00"&intv_mo, 2)            'formatting variables in to 2 digit strings - because MAXIS
	intv_day = right("00"&intv_day, 2)
	intv_yr = right(intv_yr, 2)
	intv_date_to_check = intv_mo & " " & intv_day & " " & intv_yr

	If confirm_update_prog = 1 Then     'If the dialog selects to have PROG updated
		CALL back_to_SELF               'Need to do this because we need to go to the footer month of the application and we may be in a different month

		CALL navigate_to_MAXIS_screen ("STAT", "PROG")  'Now we can navigate to PROG in the application footer month and year
		PF9                                             'Edit

		If prog_update_cash_1_checkbox = checked Then
			EMWriteScreen intv_mo, 6, 55               'CASH 1 Row
			EMWriteScreen intv_day, 6, 58
			EMWriteScreen intv_yr, 6, 61
		End If
		If prog_update_cash_2_checkbox = checked Then
			EMWriteScreen intv_mo, 7, 55               'CASh 2 Row
			EMWriteScreen intv_day, 7, 58
			EMWriteScreen intv_yr, 7, 61
		End If
		If prog_update_emer_checkbox = checked Then
			EMWriteScreen intv_mo, 8, 55               'EMER Row
			EMWriteScreen intv_day, 8, 58
			EMWriteScreen intv_yr, 8, 61
		End If
		If prog_update_grh_checkbox = checked Then
			EMWriteScreen intv_mo, 9, 55               'GRH Row
			EMWriteScreen intv_day, 9, 58
			EMWriteScreen intv_yr, 9, 61
		End If
		If prog_update_snap_checkbox = checked Then
			EMWriteScreen intv_mo, 10, 55               'SNAP Row
			EMWriteScreen intv_day, 10, 58
			EMWriteScreen intv_yr, 10, 61
		End If
		EMWriteScreen left(exp_migrant_seasonal_formworker_yn, 1), 18, 67
		transmit                                    'Saving the panel

		Call HCRE_panel_bypass
		Call back_to_SELF
		Call MAXIS_background_check
	End If

	IF confirm_update_revw = 1 Then
		original_MAXIS_month = MAXIS_footer_month
		original_MAXIS_year = MAXIS_footer_year
		cash_revw_intv_date_updated = FALSE
		snap_revw_intv_date_updated = FALSE
		If the_process_for_cash = "Renewal" AND the_process_for_snap = "Renewal" AND next_cash_revw_mo = next_snap_revw_mo AND next_cash_revw_yr = next_snap_revw_yr Then
			Call back_to_SELF
			MAXIS_footer_month = next_cash_revw_mo
			MAXIS_footer_year = next_cash_revw_yr

			Call Navigate_to_MAXIS_screen("STAT", "REVW")
			PF9
			Call create_mainframe_friendly_date(CAF_datestamp, 13, 37, "YY")
			Call create_mainframe_friendly_date(interview_date, 15, 37, "YY")

			EMReadScreen cash_revw_status_code, 1, 7, 40
			EMReadScreen snap_revw_status_code, 1, 7, 60
			If cash_revw_status_code = "N" Then EMWriteScreen "I", 7, 40
			If snap_revw_status_code = "N" Then EMWriteScreen "I", 7, 60

			attempt_count = 1
			Do
				transmit
				EMReadScreen actually_saved, 7, 24, 2
				attempt_count = attempt_count + 1
				If attempt_count = 20 Then
					PF10
					revw_panel_updated = FALSE
					Exit Do
				End If
			Loop until actually_saved = "ENTER A"

			Call back_to_SELF
			Call Navigate_to_MAXIS_screen("STAT", "REVW")

			EMReadScreen updated_intv_date, 8, 15, 37
			If IsDate(updated_intv_date) = TRUE Then
				updated_intv_date = DateAdd("d", 0, updated_intv_date)
				If updated_intv_date = interview_date Then
					cash_revw_intv_date_updated = TRUE
					snap_revw_intv_date_updated = True
				End If
			End If
		Else
			If the_process_for_cash = "Renewal" Then
				Call back_to_SELF
				MAXIS_footer_month = next_cash_revw_mo
				MAXIS_footer_year = next_cash_revw_yr

				Call Navigate_to_MAXIS_screen("STAT", "REVW")
				PF9
				Call create_mainframe_friendly_date(CAF_datestamp, 13, 37, "YY")
				Call create_mainframe_friendly_date(interview_date, 15, 37, "YY")

				EMReadScreen cash_revw_status_code, 1, 7, 40
				If cash_revw_status_code = "N" Then EMWriteScreen "I", 7, 40

				attempt_count = 1
				Do
					transmit
					EMReadScreen actually_saved, 7, 24, 2
					attempt_count = attempt_count + 1
					If attempt_count = 20 Then
						PF10
						revw_panel_updated = FALSE
						Exit Do
					End If
				Loop until actually_saved = "ENTER A"


				Call back_to_SELF
				Call Navigate_to_MAXIS_screen("STAT", "REVW")

				EMReadScreen updated_intv_date, 8, 15, 37
				If IsDate(updated_intv_date) = TRUE Then
					updated_intv_date = DateAdd("d", 0, updated_intv_date)
					If updated_intv_date = interview_date Then cash_revw_intv_date_updated = TRUE
				End If
			End If
			If the_process_for_snap = "Renewal" Then
				Call back_to_SELF
				MAXIS_footer_month = next_snap_revw_mo
				MAXIS_footer_year = next_snap_revw_yr

				Call Navigate_to_MAXIS_screen("STAT", "REVW")
				PF9
				Call create_mainframe_friendly_date(CAF_datestamp, 13, 37, "YY")
				Call create_mainframe_friendly_date(interview_date, 15, 37, "YY")

				EMReadScreen cash_revw_status_code, 1, 7, 40
				EMReadScreen snap_revw_status_code, 1, 7, 60
				If cash_revw_status_code = "N" Then EMWriteScreen "I", 7, 40
				If snap_revw_status_code = "N" Then EMWriteScreen "I", 7, 60

				attempt_count = 1
				Do
					transmit
					EMReadScreen actually_saved, 7, 24, 2
					attempt_count = attempt_count + 1
					If attempt_count = 20 Then
						PF10
						revw_panel_updated = FALSE
						Exit Do
					End If
				Loop until actually_saved = "ENTER A"

				Call back_to_SELF
				Call Navigate_to_MAXIS_screen("STAT", "REVW")

				EMReadScreen updated_intv_date, 8, 15, 37
				If IsDate(updated_intv_date) = TRUE Then
					updated_intv_date = DateAdd("d", 0, updated_intv_date)
					If updated_intv_date = interview_date Then snap_revw_intv_date_updated = TRUE
				End If
			End If
		End If

		MAXIS_footer_month = original_footer_month
		MAXIS_footer_year = original_footer_year

		fail_msg = ""
		If cash_revw_intv_date_updated = FALSE AND the_process_for_cash = "Renewal" Then fail_msg = fail_msg & vbCr & vbCr & "Interview and App date on REVW for CASH in " & next_cash_revw_mo & "/" & next_cash_revw_yr
		If snap_revw_intv_date_updated = FALSE AND the_process_for_snap = "Renewal" Then fail_msg = fail_msg & vbCr & vbCr & "Interview and App date on REVW for SNAP in " & next_snap_revw_mo & "/" & next_snap_revw_yr
		If fail_msg <> "" Then MsgBox "You have requested the script update REVW with the interview date." & vbCr & vbCr & "The script was unable to update REVW completely." & vbCr & vbCr & "FAILED:" & fail_msg & vbCr & vbCr & "The REVW panel will need to be updated manually with the interview information."
	End If
End If

If expedited_determination_completed = True Then
	If developer_mode = False Then

		txt_file_name = "expedited_determination_detail_" & MAXIS_case_number & "_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
		exp_info_file_path = t_drive &"\Eligibility Support\Assignments\Expedited Information\"  & txt_file_name
		' MsgBox exp_info_file_path

		With (CreateObject("Scripting.FileSystemObject"))

			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream

			Set objTextStream = .OpenTextFile(exp_info_file_path, ForWriting, true)

			objTextStream.WriteLine ""

			objTextStream.WriteLine "CASE NUMBER ^*^*^" & MAXIS_case_number
			objTextStream.WriteLine "WORKER NAME ^*^*^" & worker_name
			objTextStream.WriteLine "WORKER USER ID ^*^*^" & user_ID_for_validation
			objTextStream.WriteLine "CASE X NUMBER  ^*^*^" & case_pw
			CAF_datestamp_new_one = CAF_datestamp
			If IsDate(CAF_datestamp) = True Then CAF_datestamp_new_one = DateAdd("d", 0, CAF_datestamp)
			objTextStream.WriteLine "DATE OF APPLICATION ^*^*^" & CAF_datestamp_new_one
			appt_notc_sent_on_new_one = appt_notc_sent_on
			If IsDate(appt_notc_sent_on) = True Then appt_notc_sent_on_new_one = DateAdd("d", 0, appt_notc_sent_on)
			objTextStream.WriteLine "APPT NOTC SENT DATE ^*^*^" & appt_notc_sent_on_new_one
			appt_date_in_note_new_one = appt_date_in_note
			If IsDate(appt_date_in_note) = True Then appt_date_in_note_new_one = DateAdd("d", 0, appt_date_in_note)
			objTextStream.WriteLine "APPT DATE ^*^*^" & appt_date_in_note_new_one
			interview_date_new_one = interview_date
			If IsDate(interview_date) = True Then interview_date_new_one = DateAdd("d", 0, interview_date)
			objTextStream.WriteLine "DATE OF INTERVIEW ^*^*^" & interview_date_new_one
			objTextStream.WriteLine "EXPEDITED SCREENING STATUS ^*^*^" & expedited_screening
			objTextStream.WriteLine "EXPEDITED DETERMINATION STATUS ^*^*^" & is_elig_XFS
			objTextStream.WriteLine "DET INCOME ^*^*^" & determined_income
			objTextStream.WriteLine "DET ASSETS ^*^*^" & determined_assets
			objTextStream.WriteLine "DET SHEL ^*^*^" & determined_shel
			objTextStream.WriteLine "DET HEST ^*^*^" & determined_utilities
			approval_date_new_one = approval_date
			If IsDate(approval_date) = True Then approval_date_new_one = DateAdd("d", 0, approval_date)
			objTextStream.WriteLine "DATE OF APPROVAL ^*^*^" & approval_date_new_one
			snap_denial_date_new_one = snap_denial_date
			If IsDate(snap_denial_date) = True Then snap_denial_date_new_one = DateAdd("d", 0, snap_denial_date)
			objTextStream.WriteLine "SNAP DENIAL DATE ^*^*^" & snap_denial_date_new_one
			objTextStream.WriteLine "SNAP DENIAL REASON ^*^*^" & snap_denial_explain
			objTextStream.WriteLine "ID ON FILE ^*^*^" & do_we_have_applicant_id
			objTextStream.WriteLine "OUTSTATE ACTION ^*^*^" & action_due_to_out_of_state_benefits
			objTextStream.WriteLine "OUTSTATE STATE ^*^*^" & other_snap_state
			other_state_reported_benefit_end_date_new_one = other_state_reported_benefit_end_date
			If IsDate(other_state_reported_benefit_end_date) = True Then other_state_reported_benefit_end_date_new_one = DateAdd("d", 0, other_state_reported_benefit_end_date)
			objTextStream.WriteLine "OUTSTATE REPORTED END DATE ^*^*^" & other_state_reported_benefit_end_date_new_one
			objTextStream.WriteLine "OUTSTATE OPENENDED ^*^*^" & other_state_benefits_openended
			other_state_verified_benefit_end_date_new_one = other_state_verified_benefit_end_date
			If IsDate(other_state_verified_benefit_end_date) = True Then other_state_verified_benefit_end_date_new_one = DateAdd("d", 0, other_state_verified_benefit_end_date)
			objTextStream.WriteLine "OUTSTATE VERIFIED END DATE ^*^*^" & other_state_verified_benefit_end_date_new_one
			mn_elig_begin_date_new_one = mn_elig_begin_date
			If IsDate(mn_elig_begin_date) = True Then mn_elig_begin_date_new_one = DateAdd("d", 0, mn_elig_begin_date)
			objTextStream.WriteLine "MN ELIG BEGIN DATE ^*^*^" & mn_elig_begin_date_new_one
			objTextStream.WriteLine "PREV POST DELAY APP ^*^*^" & case_has_previously_postponed_verifs_that_prevent_exp_snap				'(Boolean)
			previous_CAF_datestamp_new_one = previous_CAF_datestamp
			If IsDate(previous_CAF_datestamp) = True Then previous_CAF_datestamp_new_one = DateAdd("d", 0, previous_CAF_datestamp)
			objTextStream.WriteLine "PREV POST PREV DATE OF APP ^*^*^" & previous_CAF_datestamp_new_one
			objTextStream.WriteLine "PREV POST LIST ^*^*^" & prev_verif_list
			objTextStream.WriteLine "PREV POST CURR VERIF POST ^*^*^" & curr_verifs_postponed_yn
			objTextStream.WriteLine "PREV POST ONGOING SNAP APP ^*^*^" & ongoing_snap_approved_yn
			objTextStream.WriteLine "PREV POST VERIFS RECVD ^*^*^" & prev_post_verifs_recvd_yn
			objTextStream.WriteLine "EXPLAIN APPROVAL DELAYS  ^*^*^" & delay_explanation								'(all of them)
			objTextStream.WriteLine "POSTPONED VERIFICATIONS ^*^*^" & postponed_verifs_yn
			objTextStream.WriteLine "WHAT ARE THE POSTPONED VERIFICATIONS ^*^*^" & list_postponed_verifs
			objTextStream.WriteLine "FACI DELAY ACTION ^*^*^" & delay_action_due_to_faci
			objTextStream.WriteLine "FACI DENY ^*^*^" & deny_snap_due_to_faci
			objTextStream.WriteLine "FACI NAME ^*^*^" & facility_name
			objTextStream.WriteLine "FACI INELIG SNAP ^*^*^" & snap_inelig_faci_yn
			faci_entry_date_new_one = faci_entry_date
			If IsDate(faci_entry_date) = True Then faci_entry_date_new_one = DateAdd("d", 0, faci_entry_date)
			objTextStream.WriteLine "FACI ENTRY DATE ^*^*^" & faci_entry_date_new_one
			faci_release_date_new_one = faci_release_date
			If IsDate(faci_release_date) = True Then faci_release_date_new_one = DateAdd("d", 0, faci_release_date)
			objTextStream.WriteLine "FACI RELEASE DATE ^*^*^" & faci_release_date_new_one
			objTextStream.WriteLine "FACI RELEASE IN 30 DAYS ^*^*^" & release_within_30_days_yn
			objTextStream.WriteLine "DATE OF SCRIPT RUN ^*^*^" & now
			objTextStream.WriteLine "SCRIPT RUN ^*^*^INTERVIEW"

			'Close the object so it can be opened again shortly
			objTextStream.Close

		End With

	End if

	note_calculation_detail = False
	If income_review_completed = True OR assets_review_completed = True OR shel_review_completed = True Then note_calculation_detail = True

	note_case_situation_details = False
	If action_due_to_out_of_state_benefits <> "" OR prev_post_verif_assessment_done = True OR faci_review_completed = True Then note_case_situation_details = True

	'creating a custom header: this is read by BULK - EXP SNAP REVIEW script so don't mess this please :)
	If IsDate(snap_denial_date) = TRUE Then
		case_note_header_text = "Expedited Determination: SNAP to be denied"
	Else
		IF is_elig_XFS = True then
			case_note_header_text = "Expedited Determination: SNAP appears expedited"
		ELSEIF is_elig_XFS = False then
			case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
		END IF
	End If

	'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------
	navigate_to_MAXIS_screen "CASE", "NOTE"

	Call start_a_blank_case_note
	Call write_variable_in_case_note (case_note_header_text)
	If interview_date <> "" Then Call write_variable_in_case_note (" - Interview completed on: " & interview_date & " and full Expedited Determination Done")
	IF exp_screening_note_found = TRUE Then
		Call write_variable_in_case_note ("Info from INITIAL EXPEDTIED SCREENING (resident reported on Application)")
		Call write_variable_in_case_note ("  Expedited Screening found: " & expedited_screening)
		Call write_variable_in_case_note ("  Based on: Income:  $ " & right("        " & exp_q_1_income_this_month, 8) & ", Assets:    $ " & right("        " & exp_q_2_assets_this_month, 8)    & ", Totaling: $ " & right("        " & caf_1_resources, 8))
		Call write_variable_in_case_note ("            Shelter: $ " & right("        " & exp_q_3_rent_this_month, 8)   & ", Utilities: $ " & right("        " & exp_q_4_utilities_this_month, 8) & ", Totaling: $ " & right("        " & caf_1_expenses, 8))
		Call write_variable_in_case_note ("No case action can be taken from screening alone, info may change at intrvw.")
		Call write_variable_in_case_note ("---")
	End If
	If IsDate(snap_denial_date) = TRUE Then
		Call write_variable_in_CASE_NOTE("SNAP to be denied on " & snap_denial_date & ". Since case is not SNAP eligible, case cannot receive Expedited issuance.")
		If is_elig_XFS = TRUE Then
			Call write_variable_with_indent_in_CASE_NOTE("Case is determined to meet criteria based upon income alone.")
			Call write_variable_with_indent_in_CASE_NOTE("Expedited approval requires case to be otherwise eligble for SNAP and this does not meet this criteria.")
		ElseIf is_elig_XFS = False Then
			Call write_variable_with_indent_in_CASE_NOTE("Expedited SNAP cannot be approved as case does not meet all criteria")
		End If
		Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Denial", snap_denial_explain)
	Else
		Call write_variable_in_case_note ("Info from Interview - Expedited Determination Completed:")
		IF is_elig_XFS = TRUE Then
			Call write_variable_in_case_note ("  Case is determined to meet criteria for Expedited SNAP.")
			If IsDate(approval_date) = False AND delay_explanation <> "" Then
				Call write_variable_in_case_note (" - Approval of Expedited SNAP cannot be completed due to:")
				' delay_explanation = THIS NEEDS TO BE AN ARRAY
				If InStr(delay_explanation, ";") = 0 Then
					delay_explain_array = Array(delay_explanation)
				Else
					delay_explain_array = Split(delay_explanation, ";")
				End If
				counter = 1
				For each item in delay_explain_array
					item = trim(item)
					Call write_variable_with_indent_in_CASE_NOTE(counter & ". " & item)
					counter = counter + 1
				Next
			End If
		End If
		IF is_elig_XFS = FALSE Then Call write_variable_in_case_note ("  Case does not meet Expedited SNAP criteria.")
		Call write_variable_in_case_note ("  Based on: Income:  $ " & right("        " & determined_income, 8) & ", Assets:    $ " & right("        " & determined_assets, 8)   & ", Totaling: $ " & right("        " & calculated_resources, 8))
		Call write_variable_in_case_note ("            Shelter: $ " & right("        " & determined_shel, 8)   & ", Utilities: $ " & right("        " & determined_utilities, 8) & ", Totaling: $ " & right("        " & calculated_expenses, 8))
		Call write_variable_in_CASE_NOTE("  --- Expedited Criteria Tests ---")
		If calculated_low_income_asset_test = False Then Call write_variable_in_case_note("  FAILED - Resources Less than or Equal to $100 and Income Less than $150")
		If calculated_low_income_asset_test = True Then Call write_variable_in_case_note("  PASSED - Resources Less than or Equal to $100 and Income Less than $150")
		If calculated_resources_less_than_expenses_test = False Then Call write_variable_in_case_note("  FAILED - Resources Plus Income Less than Shelter Costs")
		If calculated_resources_less_than_expenses_test = True Then Call write_variable_in_case_note("  PASSED - Resources Plus Income Less than Shelter Costs")
		Call write_variable_in_case_note ("---")
		IF is_elig_XFS = TRUE Then
			Call write_variable_in_case_note ("Important Details")
			Call write_bullet_and_variable_in_case_note ("Date of Application", CAF_datestamp)
			Call write_bullet_and_variable_in_case_note ("Date of Interview", interview_date)
			Call write_bullet_and_variable_in_case_note ("Date of Approval", approval_date)
			' Call write_bullet_and_variable_in_case_note ("Reason for Delay", delay_explanation)
			Call write_bullet_and_variable_in_CASE_NOTE("Postponed Verifs", list_postponed_verifs)
			Call write_variable_in_case_note ("---")
		End If
		If note_calculation_detail = True Then
			Call write_variable_in_case_note ("* Additional Notes about these amounts:")
			If income_review_completed = True Then
				' Call write_variable_in_case_note ("*   INCOME Details:")
				If jobs_income_yn = "Yes" Then
					' Call write_variable_in_case_note ("    - JOBS")
					for the_job = 0 to UBound(EXP_JOBS_ARRAY, 2)
						If IsNumeric(EXP_JOBS_ARRAY(jobs_wage_const, the_job)) = True AND IsNumeric(EXP_JOBS_ARRAY(jobs_hours_const, the_job)) = True Then
							Call write_variable_in_case_note ("  - JOBS: " & EXP_JOBS_ARRAY(jobs_employee_const, the_job) & " at " & EXP_JOBS_ARRAY(jobs_employer_const, the_job) & ": $" & EXP_JOBS_ARRAY(jobs_wage_const, the_job) & "/hr at " & EXP_JOBS_ARRAY(jobs_hours_const, the_job) & " hrs/wk.")
							Call write_variable_in_case_note ("            - Monthly Gross: $" & EXP_JOBS_ARRAY(jobs_monthly_pay_const, the_job))
						End If
					Next
				End If
				If busi_income_yn = "Yes" Then
					' Call write_variable_in_case_note ("    - SELF EMPLOYMENT")
					for the_busi = 0 to UBound(EXP_BUSI_ARRAY, 2)
						Call write_variable_in_case_note ("  - BUSI: " & EXP_BUSI_ARRAY(busi_owner_const, the_busi) & " for " & EXP_BUSI_ARRAY(busi_info_const, the_busi) & ".")
						Call write_variable_in_case_note ("            - Monthly Gross: $" & EXP_BUSI_ARRAY(busi_monthly_earnings_const, the_busi))
					Next
				End If
				If unea_income_yn = "Yes" Then
					' Call write_variable_in_case_note ("    - UNEARNED INCOME")
					for the_unea = 0 to UBound(EXP_UNEA_ARRAY, 2)
						Call write_variable_in_case_note ("  - UNEA: " & EXP_UNEA_ARRAY(unea_owner_const, the_unea) & " from " & EXP_UNEA_ARRAY(unea_info_const, the_unea) & ".")
						Call write_variable_in_case_note ("            - Monthly Gross: $" & EXP_UNEA_ARRAY(unea_monthly_earnings_const, the_unea))
					Next
				End If
				' app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, JOBS_ARRAY, BUSI_ARRAY, UNEA_ARRAY)
			End If
			If assets_review_completed = True Then
				' Call write_variable_in_case_note ("*   ASSET Details:")
				If cash_amount_yn = "Yes" Then Call write_variable_in_case_note ("  - CASH: Amount: $" & cash_amount)
				If bank_account_yn = "Yes" Then
					' Call write_variable_in_case_note ("    - BANK ACCOUNTS")
					For the_acct = 0 to UBound(EXP_ACCT_ARRAY, 2)
						If EXP_ACCT_ARRAY(account_type_const, the_acct) <> "Select One..." Then
							acct_info = "  - ACCT: " & EXP_ACCT_ARRAY(account_type_const, the_acct)
							If EXP_ACCT_ARRAY(bank_name_const, the_acct) <> "" Then acct_info = acct_info & " at " & EXP_ACCT_ARRAY(bank_name_const, the_acct)
							If EXP_ACCT_ARRAY(account_owner_const, the_acct) <> "" Then acct_info = acct_info & " owned by: " & EXP_ACCT_ARRAY(account_owner_const, the_acct)
							acct_info = acct_info & ". Balance: $" & EXP_ACCT_ARRAY(account_amount_const, the_acct)
							Call write_variable_in_case_note (acct_info)
						End If
					Next
				End If
				' app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, ACCOUNTS_ARRAY)
			End If
			If shel_review_completed = True Then
				' Call write_variable_in_case_note ("*   SHELTER Details:")
				first_housing_detail = True
				If rent_amount <> "" OR lot_rent_amount <> "" OR mortgage_amount <> "" OR insurance_amount <> "" OR tax_amount <> "" OR room_amount <> "" OR garage_amount <> "" Then

					Call write_variable_in_case_note ("  - SHEL: Rent:     $ " & right("    " & rent_amount, 4)    &  "   -   Lot Rent:  $" & right("    " & lot_rent_amount, 4))
					Call write_variable_in_case_note ("          Mortgage: $ " & right("    " & mortgage_amount, 4) & "   -   Insurance: $" & right("    " & insurance_amount, 4))
					Call write_variable_in_case_note ("          Tax:      $ " & right("    " & tax_amount, 4)      & "   -   Room:      $" & right("    " & room_amount, 4))
					Call write_variable_in_case_note ("          Garage:   $ " & right("    " & garage_amount, 4))
					Call write_variable_in_case_note ("          SUBSIDY:  $ " & right("    " & subsidy_amount, 4))
				End If
			End If
		End If
		' Call write_variable_in_case_note ("*   UTILITY Details:")
		If all_utilities <> "" Then Call write_variable_in_case_note ("  - HEST: " & all_utilities)
		' app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)

	End If

	If note_case_situation_details = True Then
		Call write_variable_in_case_note ("---")
		Call write_variable_in_case_note ("Additional details about this case:")

		If action_due_to_out_of_state_benefits <> "" Then Call write_variable_in_case_note ("* SNAP in Another State")
		If action_due_to_out_of_state_benefits = "DENY" Then
			Call write_variable_in_case_note ("*   SNAP to be DENIED as active in another state for the application processing 30 days.")
			If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
			Call write_variable_in_case_note ("      - Date of Application: " & CAF_datestamp)
			Call write_variable_in_case_note ("      - Day 30: " & day_30_from_application)
			If IsDate(other_state_verified_benefit_end_date) = True  Then
				Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_verified_benefit_end_date & " - this date has been confirmed")
			ElseIF IsDate(other_state_reported_benefit_end_date) = True Then
				Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_reported_benefit_end_date & " - reported")
			End If
			' Call write_variable_in_case_note ("      - Date of Application: " & CAF_datestamp)
		End If
		If action_due_to_out_of_state_benefits = "APPROVE" Then
			Call write_variable_in_case_note ("*   SNAP can be approved in MN for a later date.")
			If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
			Call write_variable_in_case_note ("      - Date of Application: " & CAF_datestamp)
			Call write_variable_in_case_note ("      - Begin Date of Eligibility in MN: " & mn_elig_begin_date)
			Call write_variable_in_case_note ("      - Day 30: " & day_30_from_application)
			If IsDate(other_state_verified_benefit_end_date) = True  Then
				Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_verified_benefit_end_date & " - this date has been confirmed")
			ElseIF IsDate(other_state_reported_benefit_end_date) = True Then
				Call write_variable_in_case_note ("      - End Date of Benefits in Other State: " & other_state_reported_benefit_end_date & " - reported")
			End If
		End If
		If action_due_to_out_of_state_benefits = "FOLLOW UP" Then
			Call write_variable_in_case_note ("*   Needs response/additional information and is causing a delay in processing")
			If other_snap_state <> "" Then Call write_variable_in_case_note ("      - Other State: " & other_snap_state)
			Call write_variable_in_case_note ("      - The end date of benefits is open-ended or unknown and needs response from the other state before we can take action on the case in MN.")
		End If
			' snap_in_another_state_detail(CAF_datestamp, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn, other_state_verified_benefit_end_date, mn_elig_begin_date, snap_denial_date, snap_denial_explain, action_due_to_out_of_state_benefits)

		If prev_post_verif_assessment_done = True Then
			Call write_variable_in_case_note ("* SNAP previously Approved with Postponed Verifciations")
			If case_has_previously_postponed_verifs_that_prevent_exp_snap = True Then
				eff_close_date = replace(previous_expedited_package, "/", "/1/")
				eff_close_date = DateAdd("m", 1, eff_close_date)
				eff_close_date = DateAdd("d", -1, eff_close_date)
				Call write_variable_in_case_note ("*   Expedited SNAP package cannot be approved due to unreceived postponed Verificactions")
				Call write_variable_in_case_note ("      - Previous application on " & previous_CAF_datestamp & " was approved as EXPEDITED with POSTPONED VERIFICATIONS.")
				Call write_variable_in_case_note ("      - This package closed on " & eff_close_date & ".")
				Call write_variable_in_case_note ("      - The postponed verifications have still not been received.")
				Call write_variable_in_case_note ("      - Previously postponed verifs: " & prev_verif_list)
				Call write_variable_in_case_note ("      - In order to approve the new Expedited Package for the current application, we would need to postpone verifications AGAIN.")
			End If
			If case_has_previously_postponed_verifs_that_prevent_exp_snap = False Then
				Call write_variable_in_case_note ("*   Though the case had previously postponed verifications, current Expedited can be approved")
				If prev_verifs_mandatory_yn = "No" Then Call write_variable_in_case_note ("      - The previous postponed verifications were not mandatory and case meet requirements for regular SNAP.")
				If curr_verifs_postponed_yn = "No" Then Call write_variable_in_case_note ("      - The current application does not require postponed verifications to be approved and case meet requirements for regular SNAP.")
				If ongoing_snap_approved_yn = "Yes" Then Call write_variable_in_case_note ("      - The case was approved for regular SNAP.")
				If prev_post_verifs_recvd_yn = "Yes" Then Call write_variable_in_case_note ("      - The previously postponed verifications have been received.")

			End If

			' previous_postponed_verifs_detail(case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done, delay_explanation, previous_CAF_datestamp, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn)
		End If
		If faci_review_completed = True Then
			If delay_action_due_to_faci = True Then
				Call write_variable_in_case_note ("* Resident is in a facility ")
				Call write_variable_in_case_note ("*  Expedited SNAP cannot be processed at this time.")
				If facility_name <> "" Then Call write_variable_in_case_note ("      - Facility Name: " & facility_name & " - an Ineligible SNAP Facility")
				If facility_name = "" Then Call write_variable_in_case_note ("      - Resident is in an Ineligible SNAP Facility")
				If IsDate(faci_entry_date) = True Then Call write_variable_in_case_note ("      - Facility Entry Date: " & faci_entry_date)
				If IsDate(faci_release_date) = True Then Call write_variable_in_case_note ("      - Release Date: " & faci_release_date)
				If release_date_unknown_checkbox = checked Then Call write_variable_in_case_note ("      - Release date is not known but is expected to be before " & day_30_from_application & ".")

			ElseIf deny_snap_due_to_faci = True Then
				Call write_variable_in_case_note ("* Resident is in a facility ")
				Call write_variable_in_case_note ("*   SNAP must be denied based on the current information.")
				If facility_name <> "" Then Call write_variable_in_case_note ("      - Facility Name: " & facility_name & " - an Ineligible SNAP Facility")
				If facility_name = "" Then Call write_variable_in_case_note ("      - Resident is in an Ineligible SNAP Facility")
				If IsDate(faci_entry_date) = True Then Call write_variable_in_case_note ("      - Facility Entry Date: " & faci_entry_date)
				If IsDate(faci_release_date) = True Then Call write_variable_in_case_note ("      - Release Date: " & faci_release_date)
				If release_date_unknown_checkbox = checked Then Call write_variable_in_case_note ("      - Release date is not known but is expected to be after " & day_30_from_application & ".")

			End If
			' household_in_a_facility_detail(delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, delay_explanation, snap_denial_explain, snap_denial_date, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn)
		End If
	End If

	Call write_variable_in_case_note ("---")
	Call write_variable_in_case_note(worker_signature)
	PF3
End If

qual_questions_yes = FALSE
If qual_question_one = "Yes" Then qual_questions_yes = TRUE
If qual_question_two = "Yes" Then qual_questions_yes = TRUE
If qual_question_three = "Yes" Then qual_questions_yes = TRUE
If qual_question_four = "Yes" Then qual_questions_yes = TRUE
If qual_question_five = "Yes" Then qual_questions_yes = TRUE

If qual_questions_yes = TRUE Then
	STATS_manualtime = STATS_manualtime + 60
	Call start_a_blank_CASE_NOTE

	Call write_variable_in_CASE_NOTE("CAF Qualifying Questions had an answer of 'YES' for at least one question")
	If qual_question_one = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fraud/DISQ for IPV (program violation)", qual_memb_one)
	If qual_question_two = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("SNAP in more than One State", qual_memb_two)
	If qual_question_three = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Fleeing Felon", qual_memb_three)
	If qual_question_four = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Drug Felony", qual_memb_four)
	If qual_question_five = "Yes" Then Call write_bullet_and_variable_in_CASE_NOTE("Parole/Probation Violation", qual_memb_five)
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	PF3
End If

Call write_verification_CASE_NOTE(create_verif_note)
call write_interview_CASE_NOTE

Call script_end_procedure_with_error_report(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/3/2024
'--Tab orders reviewed & confirmed----------------------------------------------10/3/2024
'--Mandatory fields all present & Reviewed--------------------------------------10/3/2024
'--All variables in dialog match mandatory fields-------------------------------10/3/2024
'Review dialog names for content and content fit in dialog----------------------10/3/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------10/3/2024
'--Create a button to reference instructions------------------------------------10/3/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/3/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------10/3/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------10/3/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/3/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------10/3/2024
'--PRIV Case handling reviewed -------------------------------------------------10/3/2024
'--Out-of-County handling reviewed----------------------------------------------10/3/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/3/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/3/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/3/2024
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------10/3/2024
'--Script name reviewed---------------------------------------------------------10/3/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/3/2024
'--comment Code-----------------------------------------------------------------10/3/2024
'--Update Changelog for release/update------------------------------------------10/3/2024
'--Remove testing message boxes-------------------------------------------------10/3/2024
'--Remove testing code/unnecessary code-----------------------------------------10/3/2024
'--Review/update SharePoint instructions----------------------------------------10/3/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/3/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------10/3/2024
'--Update project team/issue contact (if applicable)----------------------------10/3/2024
