'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
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
call changelog_update("02/27/2023", "Reference updated for information about EBT cards. The button to open a webpage about EBT cards has been changed to open the current page mmanaged by Accounting instead of the previous Temporary Program Changes page.", "Casey Love, Hennepin County")
call changelog_update("11/14/2022", "Created button to link to the interpreter service request.", "Casey Love, Hennepin County")
call changelog_update("10/17/2022", "Update messaging on Agency Signature on Application Forms. Hennepin County does not require an Agency Signature on any application form. ##~## ##~##See the HSR Manual for more Information - on the Applications Page. ##~##", "Casey Love, Hennepin County")
CALL changelog_update("06/21/2022", "Updated handling for non-disclosure agreement.", "MiKayla Handley, Hennepin County") '#493
call changelog_update("03/29/2022", "Removed ApplyMN as application type option.", "Ilse Ferris, Hennepin County")
call changelog_update("03/09/2022", "For MFIP Applications an MFIP Orientation to Financial Services is required and should be completed during the interview.##~## ##~##Currently, the script does not have functionality to support the details of the MFIP Orientation Requuirement. This functionality is in the process of being built and tested.##~## ##~##Until that update is complete and ready to be released, we have added a dialog with referece links to the policy requirements that will serve as a placeholder for when to complete the MFIP Orientation during the Interview.##~##", "Casey Love, Hennepin County")
call changelog_update("07/29/2021", "TESTING UPDATES##~##We made a couple changes.##~## ##~##Added a 'Worker Signature' box to the first dialog as that was missing.##~##Updated the look of the first dialog and added some guidance in pop-up boxes.##~##Changed the 'Error Message' handling in the dialog so if you have to 'BACK' to a question from the last page, it will let you.##~##Removed the 'Update PROG' functionality, since it is broken.##~## ##~##Another addition is a new tool in UTILITIES to open a PDF that was previously created. Go try it out!##~##", "Casey Love, Hennepin County")
call changelog_update("07/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS ==============================================================================================================

const ref_number					= 0
const access_denied					= 1
const full_name_const				= 2
const last_name_const				= 3
const first_name_const				= 4
const mid_initial					= 5
const other_names					= 6
const age							= 7
const date_of_birth					= 8
const ssn							= 9
const ssn_verif						= 10
const birthdate_verif				= 11
const gender						= 12
const race							= 13
const spoken_lang					= 14
const written_lang					= 15
const interpreter					= 16
const alias_yn						= 17
const ethnicity_yn					= 18
const id_verif						= 19
const rel_to_applcnt				= 20
const cash_minor					= 21
const snap_minor					= 22
const marital_status				= 23
const spouse_ref					= 24
const spouse_name					= 25
const last_grade_completed 			= 26
const citizen						= 27
const other_st_FS_end_date 			= 28
const in_mn_12_mo					= 29
const residence_verif				= 30
const mn_entry_date					= 31
const former_state					= 32
const fs_pwe						= 33
const button_one					= 34
const button_two					= 35
const imig_status 					= 36
const clt_has_sponsor				= 37
const client_verification			= 38
const client_verification_details	= 39
const client_notes					= 40
const intend_to_reside_in_mn		= 41
const race_a_checkbox				= 42
const race_b_checkbox				= 43
const race_n_checkbox				= 44
const race_p_checkbox				= 45
const race_w_checkbox				= 46
const snap_req_checkbox				= 47
const cash_req_checkbox				= 48
const emer_req_checkbox				= 49
const none_req_checkbox				= 50
const ssn_no_space					= 51
const edrs_msg						= 52
const edrs_match					= 53
const edrs_notes 					= 54
const ignore_person                 = 55
const pers_in_maxis                 = 56
const memb_is_caregiver             = 57
const cash_request_const            = 58
const hours_per_week_const          = 59
const exempt_from_ed_const          = 60
const comply_with_ed_const          = 61
const orientation_needed_const      = 62
const orientation_done_const        = 63
const orientation_exempt_const      = 64
const exemption_reason_const        = 65
const emps_exemption_code_const     = 66
const choice_form_done_const        = 67
const orientation_notes             = 68
const last_const					= 69

Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(last_const, 0)

'HERE we are declaring some information about the questions that we ask. '
' Generally:
' - question number
' - question wording
' - caf answer yes/no
' - caf write in
' - interview notes
' - verifications
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

function add_new_HH_MEMB()



end function
' show_pg_one_memb01_and_exp
' show_pg_one_address
' show_pg_memb_list
' show_q_1_6
' show_q_7_11
' show_q_14_15
' show_q_21_24
' show_qual
' show_pg_last
'
' update_addr
' update_pers

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

full_err_msg = full_err_msg & "~!~" & "1^* CAF DATESTAMP ##~##   - Enter a valid date for the CAF datestamp.##~##"

function check_for_errors(interview_questions_clear)
	' If  Then err_msg = err_msg & "~!~" & "1^* FIELD##~##   - "
	' page_display = show_pg_one_memb01_and_exp
	' If current_listing = "1"  Then tagline = ": Expedited"        'Adding a specific tagline to the header for the errors
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
	' If snap_status <> "ACTIVE" Then
	' 	intv_app_month_income = trim(intv_app_month_income)
	' 	intv_app_month_asset = trim(intv_app_month_asset)
	' 	intv_app_month_housing_expense = trim(intv_app_month_housing_expense)
	'
	' 	If intv_app_month_income = "" Then intv_app_month_income = 0
	' 	If intv_app_month_asset = "" Then intv_app_month_asset = 0
	' 	If intv_app_month_housing_expense = "" Then intv_app_month_housing_expense = 0
	'
	' 	If IsNumeric(intv_app_month_income) = False Then err_msg = err_msg & "~!~" & "1 ^* What is the total of the income received in the month of application?##~##   - Enter the amount of income in the month of application as a number. We MUST gather the income in the application month.##~##"
	' 	If IsNumeric(intv_app_month_asset) = False Then err_msg = err_msg & "~!~" & "1 ^* Use the best detail of assets the resident has available. Liquid Asset amount?##~##   - Enter the total assets in the month of application as a number. We MUST gather the assets in the application month.##~##"
	' 	If IsNumeric(intv_app_month_housing_expense) = False Then err_msg = err_msg & "~!~" & "1 ^* What is the housing expense (Rend, Mortgage, etc)##~##   - Enter the rent/mortgage in the month of application as a number. We MUST gather the expenses in the application month.##~##"
	'
	' 	'If Interview utilities have no checkmarks - then we need a checkmoark - if none - then check none
	' 	If intv_exp_pay_heat_checkbox = unchecked AND intv_exp_pay_ac_checkbox = unchecked AND intv_exp_pay_electricity_checkbox = unchecked AND intv_exp_pay_phone_checkbox = unchecked AND intv_exp_pay_none_checkbox = unchecked Then err_msg = err_msg & "~!~" & "1^* What utilities expenses exist?##~##   - You must indicate which utilities expenses the household has. If there are none, check the box for 'NONE'"
	' 	 		'If non is checked and others are checked - we need to resolve
	' 	If intv_exp_pay_none_checkbox = checked AND (intv_exp_pay_heat_checkbox = checked OR intv_exp_pay_ac_checkbox = checked OR intv_exp_pay_electricity_checkbox = checked OR intv_exp_pay_phone_checkbox = checked) Then err_msg = err_msg & "~!~" & "1^* What utilities expenses exist?##~##   - You have selected 'None' for utilities expenses and also selected one or more of the utilities. If 'None' you must not select one of utilities, but if there is utilities expense, you should not select 'None'."
	' End If


	' If current_listing = "2"  Then tagline = ": CAF ADDR"
		'If living situation is 'Blank' or 'Unknown' - ask it and update
	If living_situation = "10 - Unknown" OR living_situation = "Blank" or living_situation = "Select" Then err_msg = err_msg & "~!~" & "2 ^* Living Situation?##~##   - Clarify the living situation with the resident for entry."

	' If current_listing = "3"  Then tagline = ": CAF MEMBs"
		'If IMIG Statis is not blank - require sponsor information
		'require 'intends to reside in MN
		'ID for 01? Other caregiver?
	For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
            HH_MEMB_ARRAY(imig_status, the_memb) = trim(HH_MEMB_ARRAY(imig_status, the_memb))
    		If HH_MEMB_ARRAY(imig_status, the_memb) <> "" AND HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "" Then err_msg = err_msg & "~!~" & "3 ^* Sponsor?##~##   - Since there is immigration details listed for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ", you need to ask and record if this resident has a sponsor."
    		If HH_MEMB_ARRAY(intend_to_reside_in_mn, the_memb) = "" Then err_msg = err_msg & "~!~" & "3 ^* Intends to Reside in MN##~##   - Indicate if this resident (" & HH_MEMB_ARRAY(full_name_const, the_memb) & ") intends to reside in MN."
    		If the_memb = 0 AND (HH_MEMB_ARRAY(id_verif, the_memb) = "" OR HH_MEMB_ARRAY(id_verif, the_memb) = "NO - No Veer Prvd") Then err_msg = err_msg & "~!~" & "3 ^* Identidty Verification##~##   - Identity is required for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ". Enter the ID information on file/received or indicate that it has been requested."
        End If
	Next

	' If current_listing = "4"  Then tagline = ": Q. 1- 6"
		'if children in home - school notes need detail
	question_3_interview_notes = trim(question_3_interview_notes)
	If school_age_children_in_hh = True AND question_3_interview_notes = "" Then err_msg = err_msg & "~!~" & "4 ^* 3. Is anyone in the household attending school? Interview Notes:##~##   - Additional detail about school is needed since this household has children. Gather information about child(ren)'s grade level, district/school, and status.'##~##"

	' If current_listing = "5"  Then tagline = ": Q. 7 - 11"
		'if SNAP - must select PWE'
	If snap_status <> "INACTIVE" AND pwe_selection = "Select One..." Then err_msg = err_msg & "~!~" & "5 ^* Principal Wage Earner##~##   - Since this we have SNAP to consider, you must indicate who the resident selects as PWE."

	' If current_listing = "6"  Then tagline = ": Q. 12 - 13"

	' If current_listing = "7"  Then tagline = ": Q. 14 - 15"

	' If current_listing = "8"  Then tagline = ": Q. 16 - 20"

	' If current_listing = "9"  Then tagline = ": Q. 21 - 24"

	' If current_listing = "10" Then tagline = ": CAF QUAL Q"
		'if any question is 'Yes' Then must have a person selected
		qual_memb_one = trim(qual_memb_one)
		qual_memb_two = trim(qual_memb_two)
		qual_memb_there = trim(qual_memb_there)
		qual_memb_four = trim(qual_memb_four)
		qual_memb_five = trim(qual_memb_five)
		If qual_question_one = "?" OR (qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type")) Then
			err_msg = err_msg & "~!~" & "10^* Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			If qual_question_one = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
			If qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
		End If
		If qual_question_two = "?" OR (qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type")) Then
			err_msg = err_msg & "~!~" & "10^* Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			If qual_question_two = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
			If qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
		End If
		If qual_question_three = "?" OR (qual_question_three = "Yes" AND (qual_memb_there = "" OR qual_memb_there = "Select or Type")) Then
			err_msg = err_msg & "~!~" & "10^* Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			If qual_question_three = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
			If qual_question_three = "Yes" AND (qual_memb_there = "" OR qual_memb_there = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
		End If
		If qual_question_four = "?" OR (qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type")) Then
			err_msg = err_msg & "~!~" & "10^* Has anyone in your household been convicted of a drug felony in the past 10 years?"
			If qual_question_four = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
			If qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
		End If
		If qual_question_five = "?" OR (qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type")) Then
			err_msg = err_msg & "~!~" & "10^* Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			If qual_question_five = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
			If qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
		End If


		'THERE WILL BE MORE ONCE THE BENEFIT DETAILS ARE ENTERED

	' If current_listing = "12" Then tagline = ": Discrepancies"
		'If no phone number - confirm no phone number
		'If homeless and no mailing address - confirm and explain about mail
		'If out of county - confirm and explain transfer
		'rent on CAF1 and Q14 do not match
		'utilities on CAF1 and Q15 do not match
	'
	'
	' If current_listing = "13" Then tagline = ": Expedited"
		If expedited_determination_needed = True Then
			If expedited_determination_completed = False Then err_msg = err_msg & "~!~" & "13 ^* Expedited##~##   - You must complete the process for the Expedited Determination. Press the 'EXPEDITED' button on the right and complete all steps."
		End If
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	If err_msg = "" Then interview_questions_clear = TRUE

	If interview_questions_clear = TRUE Then
		' If current_listing = "11" Then tagline = ": CAF Last Page"
		'Both signatures - cannot be select or type or blank
		signature_detail = trim(signature_detail)
		second_signature_detail = trim(second_signature_detail)
		signature_person = trim(signature_person)
		second_signature_person = trim(second_signature_person)
		If signature_detail = "Select or Type" OR signature_detail = "" Then err_msg = err_msg & "~!~" & "11^* Signature of Primary Adult##~##   - Indicate how the signature information has been received (or not received)."
		If second_signature_detail = "Select or Type" OR second_signature_detail = "" Then err_msg = err_msg & "~!~" & "11^* Signature of Other Adult##~##   - Indicate how the second signature information has been received (or not received). If no second adult is on the case or the signature of the second adult is not required, select 'Not Required'."
		'If signatires are signed or verbal - then person and date must be completed
		If signature_detail = "Signature Completed" OR signature_detail  = "Accepted Verbally" Then
			If signature_person = "" AND signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & "11^* Signature of Primary Adult - person##~##   - Since the signature was completed, indicate whose sigature it is."
			If IsDate(signature_date) = False Then
				err_msg = err_msg & "~!~" & "11^* Signature of Primary Adult - date##~##   - Enter the date of the signature as a valid date."
			Else
				If DateDiff("d", date, signature_date) > 0 Then err_msg = err_msg & "~!~" & "11^* Signature of Primary Adult - date##~##   - The date of the primary signature cannot be in the future."
			End If
		End If
		If second_signature_detail = "Signature Completed" OR second_signature_detail  = "Accepted Verbally" Then
			If second_signature_person = "" AND second_signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & "11^* Signature of Other Adult - person##~##   - Since the secondary adult signature was completed, indicate whose sigature it is."
			If IsDate(second_signature_date) = False Then
				err_msg = err_msg & "~!~" & "11^* Signature of Other Adult - date##~##   - Enter the date of the signature as a valid date."
			Else
				If DateDiff("d", date, second_signature_date) > 0 Then err_msg = err_msg & "~!~" & "11^* Signature of Other Adult - date##~##   - The date of the primary signature cannot be in the future."
			End If
		End If
		'Interview date must be a date and not in the future
		' If  Then err_msg = err_msg & "~!~" & "11^* FIELD##~##   - "
		If IsDate(interview_date) = False Then
			err_msg = err_msg & "~!~" & "11^* Interview Date##~##   - Enter the date of the interview as a valid date."
		Else
			If DateDiff("d", date, interview_date) > 0 Then err_msg = err_msg & "~!~" & "11^* Interview Date##~##   - The date of the interview cannot be in the future."
		End If

		'If APP Date is too far away - explain delays
		'If APP Date is blank - add app date, deny date, or explain delays
		'If Deny date exists - explain denial


		' If snap_status = "PENDING" Then
		' 	If trim(snap_denial_date) <> "" AND IsDate(snap_denial_date) = FALSE Then
		' 		err_msg = err_msg & "~!~11^* SNAP DENIAL DATE ##~##   - This is a a SNAP case at application. You entered something in the SNAP denial date but it does not appear to be a date. Please list the date that SNAP will be denied if SNAP is being denied."
		' 	ElseIf IsDate(snap_denial_date) = TRUE Then
		' 		If DateDiff("d", date, snap_denial_date) > 0 Then err_msg = err_msg & "~!~11^* SNAP DENIAL DATE ##~##   - The denial date is listed as a future date. Review the date entered in the SNAP denial date field."
		' 		If trim(snap_denial_explain) = "" Then err_msg = err_msg & "~!~11^* EXPLAIN DENIAL ##~##   - Since you have a denial date listed, add some detail to explain the denial reason or other information."
		' 	ElseIf trim(snap_denial_date) = "" Then
		' 		If case_is_expedited = True Then
		' 			If IsDate(exp_snap_approval_date) = TRUE Then
		' 				If DateDiff("d", date, exp_snap_approval_date) > 0 Then
		' 					err_msg = err_msg & "~!~11^* EXP APPROVAL DATE ##~##   - The date listed in the expedited approval date is a future date. Please review the date listed and reenter if necessary."
		' 				ElseIf DateDiff("d", CAF_datestamp, exp_snap_approval_date) > 7 AND trim(exp_snap_delays) = "" Then
		' 					err_msg = err_msg & "~!~11^* EXPLAIN DELAYS ##~##   - Since Expedited SNAP is not approved within 7 days of the date of application, pease explain the reason for the delay."
		' 				End If
		' 			Else
		' 				If trim(exp_snap_delays) = "" Then err_msg = err_msg & "~!~11^* EXPLAIN DELAYS ##~##   - Since the Expedited SNAP does not have an approval date yet, either explain the reason for the delay or indicate the date of Expedited SNAP Approval."
		' 			End If
		' 		End If
		' 	End If
		' End If
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
			' If  = "" Then err_msg = err_msg & "~!~11^* TITLE ##~##   - MESSAGE"
			' If  = "" Then err_msg = err_msg & "~!~11^* TITLE ##~##   - MESSAGE"
 		End If


		If disc_no_phone_number = "EXISTS" Then err_msg = err_msg & "~!~12^* PHONE CONTACT Clarification ##~##   - Since no phone numbers were listed - confirm with the resident about phone contact and clarify."
		If disc_homeless_no_mail_addr = "EXISTS" Then err_msg = err_msg & "~!~12^* HOMELESS MAILING Clarification ##~##   - Since this case is listed as Homeless - confirm you have discussed mailing and responses."
		If disc_out_of_county = "EXISTS" Then err_msg = err_msg & "~!~12^* OUT OF COUNTY Clarification ##~##   - Since this case is indicated as being out of county - confirm you have explained case transfers."
		If disc_rent_amounts = "EXISTS" Then err_msg = err_msg & "~!~12^* HOUSING EXPENSE Clarification ##~##   - Since the amounts reported on the CAF for Housing Expense appear to have a discrepancy - clarify which is accurate."
		If disc_utility_amounts = "EXISTS" Then err_msg = err_msg & "~!~12^* UTILITY EXPENSE Clarification ##~##   - Since the amounts reported on the CAF for Utility Expense appear to have a discrepancy - clarify which is accurate."
		' If  = "" Then err_msg = err_msg & "~!~12^* TITLE ##~##   - MESSAGE"


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
		    Text 35, 210, 270, 10, "1. How much income (cash or checkes) did or will your household get this month?"
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
			' GroupBox 5, 200, 475, 160, "Expedited Determination"
		    ' Text 15, 210, 190, 10, "Confirm the Income received in the application month. "
		    ' Text 20, 220, 230, 10, "What is the total of the income recevied in the month of application?"
		    ' EditBox 250, 215, 55, 15, intv_app_month_income
		    ' PushButton 320, 215, 145, 15, "Resident is unsure of App Month Income", exp_income_guidance_btn
		    ' Text 15, 240, 115, 10, "Confirm the Assets the resident has."
		    ' Text 20, 250, 245, 10, "Use the best detail of assets the resident has available. Liquid Asset amount?"
		    ' EditBox 270, 245, 50, 15, intv_app_month_asset
		    ' Text 15, 270, 195, 10, "Confirm Expenses the resident has in the application month."
		    ' Text 20, 280, 180, 10, "What is the housing expense? (Rent, Mortgage, ectc.)"
		    ' EditBox 210, 275, 50, 15, intv_app_month_housing_expense
		    ' Text 20, 295, 115, 10, "What utilities expenses exist?"
		    ' CheckBox 130, 295, 30, 10, "Heat", intv_exp_pay_heat_checkbox
		    ' CheckBox 165, 295, 65, 10, "Air Conditioning", intv_exp_pay_ac_checkbox
		    ' CheckBox 235, 295, 45, 10, "Electricity", intv_exp_pay_electricity_checkbox
		    ' CheckBox 285, 295, 35, 10, "Phone", intv_exp_pay_phone_checkbox
		    ' CheckBox 330, 295, 35, 10, "None", intv_exp_pay_none_checkbox
		    ' Text 15, 315, 105, 10, "Do we have an ID verification?"
		    ' DropListBox 125, 310, 45, 45, "?"+chr(9)+"No"+chr(9)+"Yes", id_verif_on_file
		    ' Text 195, 315, 165, 10, "Check ECF, SOL-Q, and check in with the resident."
		    ' Text 15, 330, 240, 10, "Is the household active SNAP in another state for the application month?"
		    ' DropListBox 255, 325, 45, 45, "?"+chr(9)+"No"+chr(9)+"Yes", snap_active_in_other_state
		    ' Text 15, 345, 270, 10, "Was the last SNAP benefit for this case 'Expedited' with postponed verifications?"
		    ' DropListBox 285, 340, 45, 45, "?"+chr(9)+"No"+chr(9)+"Yes", last_snap_was_exp
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

						' CheckBox 330, 165, 30, 10, "Asian", HH_MEMB_ARRAY(selected_memb).race_a_checkbox
						' CheckBox 330, 175, 30, 10, "Black", HH_MEMB_ARRAY(selected_memb).race_b_checkbox
						' CheckBox 330, 185, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(selected_memb).race_n_checkbox
						' CheckBox 330, 195, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(selected_memb).race_p_checkbox
						' CheckBox 330, 205, 130, 10, "White", HH_MEMB_ARRAY(selected_memb).race_w_checkbox
						' CheckBox 70, 200, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(selected_memb).snap_req_checkbox
						' CheckBox 125, 200, 65, 10, "Cash programs", HH_MEMB_ARRAY(selected_memb).cash_req_checkbox
						' CheckBox 195, 200, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(selected_memb).emer_req_checkbox
						' CheckBox 280, 200, 30, 10, "NONE", HH_MEMB_ARRAY(selected_memb).none_req_checkbox
						' DropListBox 15, 230, 80, 45, "Yes"+chr(9)+"No", HH_MEMB_ARRAY(selected_memb).intend_to_reside_in_mn
						' EditBox 100, 230, 205, 15, HH_MEMB_ARRAY(selected_memb).imig_status
						' DropListBox 310, 230, 55, 45, "No"+chr(9)+"Yes", HH_MEMB_ARRAY(selected_memb).clt_has_sponsor
						' DropListBox 15, 260, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(selected_memb).client_verification
						' EditBox 100, 260, 435, 15, HH_MEMB_ARRAY(selected_memb).client_verification_details
						' EditBox 15, 290, 350, 15, HH_MEMB_ARRAY(selected_memb).client_notes
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
			DropListBox 70, 295, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(client_verification, selected_memb)
			EditBox 155, 295, 320, 15, HH_MEMB_ARRAY(client_verification_details, selected_memb)
			EditBox 70, 325, 405, 15, HH_MEMB_ARRAY(client_notes, selected_memb)
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
			Text 70, 285, 50, 10, "Verification"
			Text 155, 285, 65, 10, "Verification Details"
			Text 70, 315, 50, 10, "Notes:"
		ElseIf page_display = show_q_1_6 Then
			Text 510, 62, 60, 10, "Q. 1 - 6"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "1. Does everyone in your household buy, fix or eat food with you?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_1_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_1_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_1_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_1_notes
				Text 360, y_pos, 110, 10, "Q1 - Verification - " & question_1_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_1_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_1_btn
			y_pos = y_pos + 20

			'TESTING CODE
			'NEW QUESTION LAYOUT OPTION
			' Text 400, y_pos, 40, 10, "CAF Answer"
			' DropListBox 440, y_pos - 5, 35, 45, question_answers, question_1_yn
			' PushButton 400, y_pos + 10, 75, 10, "ADD WRITE-IN", question_1_add_wwrite_in_bnt
			' y_pos = y_pos + 20
			' Text 15, y_pos, 60, 10, "Verbal Answer"
			' DropListBox 75, y_pos - 5, 35, 45, question_answers, question_1_verbal_yn
			' If question_1_verif_yn <> "" Then Text 360, y_pos, 110, 10, "Q1 - Verification - " & question_1_verif_yn
			' PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_1_btn
			'
			' y_pos = y_pos + 20
			' Text 15, y_pos, 60, 10, "Interview Notes:"
			' EditBox 75, y_pos - 5, 400, 15, question_1_interview_notes
			' y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			y_pos = y_pos + 20
			' Text 20, 55, 115, 10, "buy or fix food due to a disability?"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_2_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_2_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_2_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_2_notes
				Text 360, y_pos, 110, 10, "Q2 - Verification - " & question_2_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_2_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_2_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "3. Is anyone in the household attending school?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_3_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_3_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_3_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_3_notes
				Text 360, y_pos, 110, 10, "Q3 - Verification - " & question_3_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_3_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_3_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "4. Is anyone in your household temporarily not living in your home? (eg. vacation, foster care, treatment, hospital, job search)"
			y_pos = y_pos + 20
			' Text 20, 135, 230, 10, "(for example: vacation, foster care, treatment, hospital, job search)"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_4_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_4_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_4_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_4_notes
				Text 360, y_pos, 110, 10, "Q4 - Verification - " & question_4_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_4_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_4_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			y_pos = y_pos + 20
			' Text 20, 180, 185, 10, " that limits the ability to work or perform daily activities?"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_5_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_5_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_5_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_5_notes
				Text 360, y_pos, 110, 10, "Q5 - Verification - " & question_5_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_5_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_5_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "6. Is anyone unable to work for reasons other than illness or disability?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_6_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_6_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_6_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_6_notes
				Text 360, y_pos, 110, 10, "Q6 - Verification - " & question_6_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_6_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_6_btn
			y_pos = y_pos + 20

		ElseIf page_display = show_q_7_11 Then
			Text 508, 77, 60, 10, "Q. 7 - 11"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
			' Text 20, 315, 350, 10, "- Stop working or quit a job?   - Refuse a job offer? - Ask to work fewer hours?   - Go on strike?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_7_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_7_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_7_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_7_notes
				Text 360, y_pos, 110, 10, "Q7 - Verification - " & question_7_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_7_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_7_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 65, "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_8_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_8_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_8_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_8_notes
				Text 360, y_pos, 110, 10, "Q8 - Verification - " & question_8_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 400, 10, "a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?       CAF Answer"
			DropListBox 415, y_pos - 5, 35, 45, question_answers, question_8a_yn
			y_pos = y_pos + 15
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_8_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_8_btn
			y_pos = y_pos + 25

			grp_len = 35
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then grp_len = grp_len + 20
			next
			GroupBox 5, y_pos, 475, grp_len, "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			PushButton 425, y_pos, 55, 10, "ADD JOB", add_job_btn
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_9_yn
			Text 95, y_pos, 25, 10, "write-in:"
			EditBox 120, y_pos - 5, 350, 15, question_9_notes
			' Text 360, y_pos, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			' y_pos = y_pos + 20

			' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			' y_pos = 110
			' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			First_job = TRUE
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
					If First_job = TRUE Then y_pos = y_pos + 20
					First_job = FALSE
					If JOBS_ARRAY(verif_yn, each_job) = "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					If JOBS_ARRAY(verif_yn, each_job) <> "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) & "   - Verification - " & JOBS_ARRAY(verif_yn, each_job)
					PushButton 450, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					y_pos = y_pos + 10
				End If
			next
			If First_job = TRUE Then y_pos = y_pos + 10
			y_pos = y_pos + 15

			GroupBox 5, y_pos, 475, 55, "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_10_yn
			Text 95, y_pos, 50, 10, "Gross Earnings:"
			EditBox 145, y_pos - 5, 35, 15, question_10_monthly_earnings
			Text 180, y_pos, 25, 10, "write-in:"
			If question_10_verif_yn = "" Then
				EditBox 205, y_pos - 5, 270, 15, question_10_notes
			Else
				EditBox 205, y_pos - 5, 150, 15, question_10_notes
				Text 360, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_10_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_10_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "11. Do you expect any changes in income, expenses or work hours?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_11_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_11_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_11_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_11_notes
				Text 360, y_pos, 110, 10, "Q11 - Verification - " & question_11_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_11_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_11_btn
			y_pos = y_pos + 25

			Text 5, y_pos, 75, 10, "Pricipal Wage Earner"
			DropListBox 85, y_pos - 5, 175, 45, pick_a_client, pwe_selection
			y_pos = y_pos + 10


		ElseIf page_display = show_q_12_13 Then
			Text 505, 92, 60, 10, "Q. 12 - 13"
			y_pos = 10

			GroupBox 5, y_pos, 475, 125, "12. Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			' y_pos = y_pos + 15
			PushButton 385, y_pos + 5, 90, 13, "ALL Q. 12 Answered 'No'", q_12_all_no_btn
			y_pos = y_pos + 20
			col_1_1 = 15
			col_1_2 = 55
			col_1_3 = 115

			col_2_1 = 165
			col_2_2 = 205
			col_2_3 = 260

			col_3_1 = 320
			col_3_2 = 360
			col_3_3 = 430

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_1_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_3_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_3_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_rsdi_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "RSDI                  $"
			EditBox 		col_1_3,	y_pos, 		35, 15, question_12_rsdi_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_ssi_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "SSI                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, question_12_ssi_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_va_yn
			Text 			col_3_2, 	y_pos + 5, 	70, 10, "VA                          $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_va_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_ui_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "UI                       $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, question_12_ui_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_wc_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "WC                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, question_12_wc_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_ret_yn
			Text 			col_3_2, 	y_pos + 5, 	85, 10, "Retirement Ben.     $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_ret_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_trib_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "Tribal Payments  $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, question_12_trib_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_cs_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "CSES             $"
			EditBox 		col_2_3,	y_pos, 		35, 15, question_12_cs_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_other_yn
			Text 			col_3_2, 	y_pos + 5, 	110, 10, "Other unearned       $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_other_amt
			y_pos = y_pos + 25

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_12_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_12_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_12_notes
				Text 360, y_pos, 110, 10, "Q12 - Verification - " & question_12_verif_yn
			End If
			' Text 360, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_12_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_12_btn
			y_pos = y_pos + 25

			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			GroupBox 5, y_pos, 475, 55, "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_13_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_13_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_13_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_13_notes
				Text 360, y_pos, 110, 10, "Q13 - Verification - " & question_13_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_13_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_13_btn
			y_pos = y_pos + 20

		ElseIf page_display = show_q_14_15 Then
			Text 505, 107, 60, 10, "Q. 14 - 15"
			y_pos = 10

			GroupBox 5, 10, 475, 130, "14. Does your household have the following housing expenses?"
			PushButton 385, 15, 90, 13, "ALL Q. 14 Answered 'No'", q_14_all_no_btn

			y_pos = y_pos + 15
			col_1_1 = 15
			col_1_2 = 85
			col_2_1 = 220
			col_2_2 = 290

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_rent_yn
			Text 			col_1_2, y_pos, 	70, 10, "Rent"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_subsidy_yn
			Text 			col_2_2, y_pos, 	100, 10, "Rent or Section 8 Subsidy"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_mortgage_yn
			Text 			col_1_2, y_pos, 	125, 10, "Mortgage/contract for deed payment"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_association_yn
			Text 			col_2_2, y_pos, 	70, 10, "Association fees"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_insurance_yn
			Text 			col_1_2, y_pos, 	85, 10, "Homeowner's insurance"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_room_yn
			Text 			col_2_2, y_pos, 	70, 10, "Room and/or board"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_taxes_yn
			Text 			col_1_2, y_pos, 	100, 10, "Real estate taxes"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_14_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_14_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_14_notes
				Text 360, y_pos, 110, 10, "Q14 - Verification - " & question_14_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_14_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_14_btn
			y_pos = y_pos + 25


			GroupBox 5, y_pos, 475, 135, "15. Does your household have the following utility expenses any time during the year? "
			y_pos = y_pos + 15

			col_1_1 = 20
			col_1_2 = 65

			col_2_1 = 185
			col_2_2 = 230

			col_3_1 = 335
			col_3_2 = 380

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_3_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_heat_ac_yn
			Text 			col_1_2, y_pos, 	85, 10, "Heating/air conditioning"
			DropListBox 	col_2_1, y_pos - 5, 35, 45, question_answers, question_15_electricity_yn
			Text 			col_2_2, y_pos, 	70, 10, "Electricity"
			DropListBox 	col_3_1, y_pos - 5, 35, 45, question_answers, question_15_cooking_fuel_yn
			Text 			col_3_2, y_pos, 	70, 10, "Cooking fuel"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_water_and_sewer_yn
			Text 			col_1_2, y_pos, 	75, 10, "Water and sewer"
			DropListBox 	col_2_1, y_pos - 5, 35, 45, question_answers, question_15_garbage_yn
			Text 			col_2_2, y_pos, 	60, 10, "Garbage removal"
			DropListBox 	col_3_1, y_pos - 5, 35, 45, question_answers, question_15_phone_yn
			Text 			col_3_2, y_pos, 	70, 10, "Phone/cell phone"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_liheap_yn
			Text 			col_1_2, y_pos, 375, 10, "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_15_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_15_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_15_notes
				Text 360, y_pos, 110, 10, "Q15 - Verification - " & question_15_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_15_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_15_btn
			y_pos = y_pos + 20

			Text 15, y_pos, 100, 10, "Does phone have an expense?"
			ComboBox 115, y_pos - 5, 360, 15, "Select or Type"+chr(9)+"Yes there is a cost, the bill is the responsibility of a unit member."+chr(9)+"Yes there is a cost, the household has a partial subsidy but pays a portion of the bill."+chr(9)+"No Expense, this is from a free phone program and does not cost the household anything."+chr(9)+"Yes there is a cost, optional service add-ons to a free phone program are paid by the household."+chr(9)+"No Expense, this household does not have a phone of their own."+chr(9)+question_15_phone_details, question_15_phone_details

		ElseIf page_display = show_q_16_20 Then
			Text 505, 122, 60, 10, "Q. 16 - 20"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			' Text 95, 200, 125, 10, "looking for work or going to school?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_16_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_16_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_16_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_16_notes
				Text 		360, y_pos, 	110, 10, "Q16 - Verification - " & question_16_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_16_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_16_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "17. Does anyone have costs for care of an ill/disabled adult because you or they are working, looking for work or going to school?"
			' Text 95, 245, 125, 10, "looking for work or going to school?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_17_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_17_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_17_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_17_notes
				Text 		360, y_pos, 	110, 10, "Q17 - Verification - " & question_17_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_17_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_17_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "18. Does anyone in the household pay support, or contribute to a tax dependent who does not live in your home?"
			' Text 95, 290, 215, 10, "or contribute to a tax dependent who does not live in your home?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_18_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_18_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_18_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_18_notes
				Text 		360, y_pos, 	110, 10, "Q18 - Verification - " & question_18_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_18_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_18_btn
			y_pos = y_pos + 20

			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			GroupBox 5, y_pos, 475, 55, "19. For SNAP only: Does anyone in the household have medical expenses? "
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_19_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_19_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_19_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_19_notes
				Text 		360, y_pos, 	110, 10, "Q19 - Verification - " & question_19_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 60, 	10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_19_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_19_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 100, "20. Does anyone in the household own, or is anyone buying, any of the following?"
			y_pos = y_pos + 10
			col_1_1 = 25
			col_1_2 = 90
			col_2_1 = 230
			col_2_2 = 295

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_20_cash_yn
			Text 			col_1_2, y_pos, 	70, 10, "Cash"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_20_acct_yn
			Text 			col_2_2, y_pos, 	175, 10, "Bank accounts (savings, checking, debit card, etc.)"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_20_secu_yn
			Text 			col_1_2, y_pos, 	125, 10, "Stocks, bonds, annuities, 401k, etc."
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_20_cars_yn
			Text 			col_2_2, y_pos, 	180, 10, "Vehicles (cars, trucks, motorcycles, campers, trailers)"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_20_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_20_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_20_notes
				Text 360, y_pos, 110, 10, "Q20 - Verification - " & question_20_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_20_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_20_btn
			y_pos = y_pos + 25

		ElseIf page_display = show_q_21_24 Then
			Text 505, 137, 60, 10, "Q. 21 - 24"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? "
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_21_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_21_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_21_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_21_notes
				Text 360, y_pos, 110, 10, "Q21 - Verification - " & question_21_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_21_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_21_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 55, "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_22_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_22_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_22_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_22_notes
				Text 360, y_pos, 110, 10, "Q22 - Verification - " & question_22_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_22_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_22_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 55, "23. For children under the age of 19, are both parents living in the home?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_23_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_23_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_23_notes
			Else
	 			EditBox 120, y_pos - 5, 235, 15, question_23_notes
				Text 360, y_pos, 110, 10, "Q23 - Verification - " & question_23_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_23_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_23_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 100, "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			y_pos = y_pos + 10

			col_1_1 = 25
			col_1_2 = 90
			col_2_1 = 230
			col_2_2 = 295

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox col_1_1, y_pos - 5, 60, 45, question_answers, question_24_rep_payee_yn
			Text 		col_1_2, y_pos, 	95, 10, "Representative Payee fees"
			DropListBox col_2_1, y_pos - 5, 60, 45, question_answers, question_24_guardian_fees_yn
			Text 		col_2_2, y_pos, 	105, 10, "Guardian Conservator fees"
			y_pos = y_pos + 15

			DropListBox col_1_1, y_pos - 5, 60, 45, question_answers, question_24_special_diet_yn
			Text 		col_1_2, y_pos, 	125, 10, "Physician-perscribed special diet"
			DropListBox col_2_1, y_pos - 5, 60, 45, question_answers, question_24_high_housing_yn
			Text 		col_2_2, y_pos, 	105, 10, "High housing costs"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_24_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_24_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_24_notes
				Text 360, y_pos, 110, 10, "Q24 - Verification - " & question_24_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_24_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_24_btn

		ElseIf page_display = show_qual Then
			Text 500, 152, 60, 10, "CAF QUAL Q"

			DropListBox 220, 40, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_one
			ComboBox 340, 40, 105, 45, all_the_clients, qual_memb_one
			DropListBox 220, 80, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_two
			ComboBox 340, 80, 105, 45, all_the_clients, qual_memb_two
			DropListBox 220, 110, 30, 45, "?"+chr(9)+"No"+chr(9)+"Yes", qual_question_three
			ComboBox 340, 110, 105, 45, all_the_clients, qual_memb_there
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
			Text 498, 167, 60, 10, "CAF Last Page"

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

					Text 15, y_pos, 450, 10, case_assesment_text
					y_pos = y_pos + 10

					Text 20, y_pos, 450, 20, next_steps_one
					y_pos = y_pos + 20
					Text 20, y_pos, 450, 20, next_steps_two
					y_pos = y_pos + 20
					Text 20, y_pos, 450, 20, next_steps_three
					y_pos = y_pos + 20
					Text 20, y_pos, 450, 20, next_steps_four
					y_pos = y_pos + 20
				End If
			End If
			' If snap_status = "ACTIVE" Then
			' 	Text 15, y_pos, 450, 10, "SNAP is active on this case - Expedited Determination not needed."
			' 	y_pos = y_pos + 15
			' Else
			' 	If case_is_expedited = True Then Text 15, y_pos, 325, 10, "Case appears to meet Expedited Criteria and needs to be processed using Expedited Standards."
			' 	If case_is_expedited = False Then Text 15, y_pos, 325, 10, "Case does not appear to be expedited, if that seems incorrect - review EXP Quesitons."
			' 	Text 350, y_pos, 120, 10, "CAF Date: " & CAF_datestamp
			' 	y_pos = y_pos + 10
			'
			' 	Text 25, y_pos, 120, 10, "App Month - Income: $" & intv_app_month_income
			' 	Text 150, y_pos, 75, 10, "Assets: $" & intv_app_month_asset
			' 	Text 225, y_pos, 75, 10, "Expenses: $" & app_month_expenses
			' 	y_pos = y_pos + 20
			'
			' 	If snap_status = "PENDING" Then
			' 		Text 20, y_pos, 65, 10, "EXP Approval Date:"
			' 		EditBox 90, y_pos - 5, 35, 15, exp_snap_approval_date
			' 		Text 135, y_pos, 55, 10, "Explain Delays:"
			' 		EditBox 190, y_pos - 5, 275, 15, exp_snap_delays
			' 		y_pos = y_pos + 20
			' 		Text 20, y_pos, 75, 10, "SNAP Denial Date:"
			' 		EditBox 90, y_pos - 5, 35, 15, snap_denial_date
			' 		Text 135, y_pos, 55, 10, "Explain denial:"
			' 		EditBox 190, y_pos - 5, 275, 15, snap_denial_explain
			' 		y_pos = y_pos + 20
			'
			' 	ElseIf snap_status = "INACTIVE" Then
			' 		Text 25, y_pos, 90, 10, "Review case, should SNAP be pended?"
			' 		DropListBox 115, y_pos - 5, 75, 45, "?"+chr(9)+"Yes"+chr(9)+"No", pend_snap_on_case
			' 		y_pos = y_pos + 20
			'
			' 	End If
			' 	Text 15, y_pos, 400, 10, "(Income, Assets, and Expenses are determined on the 'Expedited' page of this dialog.)"
			' 	y_pos = y_pos + 15
			' End If

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
			' expedited_info_does_not_match
			' mismatch_explanation

			' Call determine_program_and_case_status_from_CASE_CURR(
			' case_active
			' case_pending
			' case_rein
			' family_cash_case
			' mfip_case
			' dwp_case
			' adult_cash_case
			' ga_case
			' msa_case
			' grh_case
			' snap_case
			' ma_case
			' msp_case
			' unknown_cash_pending
			' unknown_hc_pending
			' ga_status
			' msa_status
			' mfip_status
			' dwp_status
			' grh_status
			' snap_status
			' ma_status
			' msp_status

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
			DropListBox 95, 290, 175, 15, "Select One..."+chr(9)+"AREP authorized verbal"+chr(9)+"AREP Authorized by entry on the CAF"+chr(9)+"AREP authorized by seperate writen document"+chr(9)+"AREP previously entered - authorization unknown"+chr(9)+"DO NOT AUTHORIZE AN AREP"+chr(9)+arep_authorization, arep_authorization
			PushButton 395, 292, 85, 13, "Save AREP Detail", save_information_btn

		ElseIf page_display = discrepancy_questions Then
			btn_pos = 180
			Text 504, btn_pos + 2, 60, 10, "Clarifications"

			y_pos = 10
			If disc_no_phone_number = "EXISTS" OR disc_no_phone_number = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 35, "No Phone Number, Review Phone Contact"
				Text 20, y_pos + 20, 165, 10, "Confirm with the resident about phone contact."
				ComboBox 185, y_pos + 15, 270, 45, "Select or Type"+chr(9)+"Confirmed No good phone contact"+chr(9)+"Added a Message Only Number"+chr(9)+"Added a Phone Number"+chr(9)+"Resident will Contact with a Phone Number once Obtained"+chr(9)+disc_phone_confirmation, disc_phone_confirmation
				y_pos = y_pos + 40
			End If
			If disc_yes_phone_no_expense = "EXISTS" OR disc_yes_phone_no_expense = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 35, "Phone Number listed, NO Phone Expense"
				Text 20, y_pos + 20, 100, 10, "Clarify how phone is paid:"
				ComboBox 120, y_pos + 15, 335, 45, "Select or Type"+chr(9)+"Phone paid by Government Free Phone Program with no expense."+chr(9)+"Phone is paid by someone out of the home, billed directly to them."+chr(9)+"Phone is a community line available for messages only."+chr(9)+"Phone is a community line in the building/residence the resident stays at."+chr(9)+disc_yes_phone_no_expense_confirmation, disc_yes_phone_no_expense_confirmation
				y_pos = y_pos + 40
			End If
			If disc_no_phone_yes_expense = "EXISTS" OR disc_no_phone_yes_expense = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 35, "No Phone Number Listed, Phone Expense Indicated"
				Text 20, y_pos + 20, 165, 10, "Clarify a phone number or explain expense:"
				ComboBox 185, y_pos + 15, 270, 45, "Select or Type"+chr(9)+"Paying phone for somone outside the home."+chr(9)+"Lost phone, number is changing."+chr(9)+"Getting a new number."+chr(9)+disc_no_phone_yes_expense_confirmation, disc_no_phone_yes_expense_confirmation
				y_pos = y_pos + 40
			End If
			If disc_homeless_no_mail_addr = "EXISTS" OR disc_homeless_no_mail_addr = "RESOLVED" Then
				grp_len = 80
				If mail_addr_street_full <> "" Then grp_len = 95
				GroupBox 10, y_pos, 455, grp_len, "Homeless, Review Mailing Options"
				Text 20, y_pos + 10, 435, 40, "Explain that actions on the case are going to come officially through the mail. General Delivery can work as a mail option, but you need to collect your mail very regularly, at least once a week, to ensure you get your informaiton and notifications timely. If you have a trusted address you can use as a mailing address, maybe a friend or family member, that is often easier to navigate. Know that much of our mail must be responded to right away, we may need to receive verification within days of a mailing."
				Text 25, y_pos + 45, 400, 10, "RESIDENCE ADDR: " & resi_addr_street_full & " " & resi_addr_city & ", " & left(resi_addr_state, 2) & " " & resi_addr_zip
				y_pos = y_pos + 65
				If mail_addr_street_full <> "" Then
					Text 25, y_pos - 5, 400, 10, "MAILING ADDR: " & mail_addr_street_full & " " & mail_addr_city & ", " & left(mail_addr_state, 2) & " " & mail_addr_zip
					y_pos = y_pos + 15
				End If
				' y_pos = y_pos + 5
				Text 20, Y_pos, 200, 10, "Confirm you have discussed the difficulties/issues with mail"
				ComboBox 210, Y_pos - 5, 245, 10, "Select or Type"+chr(9)+"Confirmed Understanding of General Delivery"+chr(9)+"Added a Trusted Mailing Address"+chr(9)+"Resident will look for a new Solution and Communicate"+chr(9)+disc_homeless_confirmation,disc_homeless_confirmation
				y_pos = y_pos + 20
			End If
			If disc_out_of_county = "EXISTS" OR disc_out_of_county = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 35, "Residence is Out of County. Review Case Transfer"
				PushButton 305, y_pos - 2, 150, 13, "HSR Manual - Transfer to Another County", open_hsr_manual_transfer_page_btn
				Text 20, y_pos + 20, 150, 10, "Confirm Out of County process discussed:"
				ComboBox 165, y_pos + 15, 290, 45, "Select or Type"+chr(9)+"Discussion Completed"+chr(9)+"County of Residence Updated"+chr(9)+disc_out_of_county_confirmation, disc_out_of_county_confirmation
				y_pos = y_pos + 40

			End If
			If disc_rent_amounts = "EXISTS" OR disc_rent_amounts = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 65, "CAF Answers for Housing Expense do not Match, Review and Clarify"
				Text 20, y_pos + 15, 400, 10, "CAF Page 1 Housing Expense: " & exp_q_3_rent_this_month
				Text 20, y_pos + 30, 400, 10, "Question 14 Housing Expense: " & question_14_summary

				Text 20, y_pos + 50, 110, 10, "Confirm Housing Expense Detail: "
				ComboBox 125, y_pos + 45, 330, 45, "Select or Type"+chr(9)+"Houshold DOES have Housing Expense"+chr(9)+"Household has NO Housing expense"+chr(9)+"Houshold has an ongoing Housing Expense but NONE in the Application month"+chr(9)+"Houshold has Housing Expense in the application months but NONE ongoing"+chr(9)+disc_rent_amounts_confirmation, disc_rent_amounts_confirmation
				y_pos = y_pos + 70
			End If
			If disc_utility_amounts = "EXISTS" OR disc_utility_amounts = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 65, "CAF Answers for Utility Expense do not Match, Review and Clarify"
				Text 20, y_pos + 15, 400, 10, "CAF Page 1 Utility Expense: " & disc_utility_caf_1_summary
				Text 20, y_pos + 30, 400, 10, "Question 15 Utility Expense: " & disc_utility_q_15_summary

				Text 20, y_pos + 50, 110, 10, "Confirm Utility Expense Detail: "
				ComboBox 125, y_pos + 45, 330, 45, "Select or Type"+chr(9)+"Household pays for Heat"+chr(9)+"Household pays for AC"+chr(9)+"Houshold pays Electricity which INCLUDES AC"+chr(9)+"Houshold pays Electricity which INCLUDES Heat"+chr(9)+"Houshold pays Electricity which INCLUDES AC and Heat"+chr(9)+"Houshold pays Electricity, but this does not include Heat or AC"+chr(9)+"Houshold pays Electricity and Phone"+chr(9)+"Houshold pays Phone Only"+chr(9)+"Houshold pays NO Utility Expenses"+chr(9)+disc_utility_amounts_confirmation, disc_utility_amounts_confirmation
				y_pos = y_pos + 70
			End If

		ElseIf page_display = expedited_determination Then
			btn_pos = 180
			If discrepancies_exist = True Then btn_pos = btn_pos + 15
			Text 505, btn_pos+2, 60, 10, "EXPEDITED"

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
		Text 485, 122, 10, 10, "8"
		Text 485, 137, 10, 10, "9"
		Text 485, 152, 10, 10, "10"
		Text 485, 167, 10, 10, "11"
		If page_display <> show_pg_one_memb01_and_exp 	Then PushButton 495, 15, 55, 13, "INTVW / CAF 1", caf_page_one_btn
		If page_display <> show_pg_one_address 			Then PushButton 495, 30, 55, 13, "CAF ADDR", caf_addr_btn
		' If page_display <> show_pg_memb_list AND page_display <> show_pg_memb_info AND  page_display <> show_pg_imig Then PushButton 485, 25, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_pg_memb_list 			Then PushButton 495, 45, 55, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_q_1_6 					Then PushButton 495, 60, 55, 13, "Q. 1 - 6", caf_q_1_6_btn
		If page_display <> show_q_7_11 					Then PushButton 495, 75, 55, 13, "Q. 7 - 11", caf_q_7_11_btn
		If page_display <> show_q_12_13 				Then PushButton 495, 90, 55, 13, "Q. 12 - 13", caf_q_12_13_btn
		If page_display <> show_q_14_15 				Then PushButton 495, 105, 55, 13, "Q. 14 - 15", caf_q_14_15_btn
		If page_display <> show_q_16_20 				Then PushButton 495, 120, 55, 13, "Q. 16 - 20", caf_q_16_20_btn
		If page_display <> show_q_21_24 				Then PushButton 495, 135, 55, 13, "Q. 21 - 24", caf_q_21_24_btn

		If page_display <> show_qual 					Then PushButton 495, 150, 55, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last 				Then PushButton 495, 165, 55, 13, "CAF Last Page", caf_last_page_btn
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
		PushButton 10, 365, 130, 15, "Interview Ended - INCOMPLETE", incomplete_interview_btn
		PushButton 140, 365, 130, 15, "View Verifications", verif_button
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn

	EndDialog

end function

function dialog_movement()
	' case_has_imig = FALSE
	' MsgBox ButtonPressed
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
	If ButtonPressed = add_verif_1_btn Then Call verif_details_dlg(1)
	If ButtonPressed = add_verif_2_btn Then Call verif_details_dlg(2)
	If ButtonPressed = add_verif_3_btn Then Call verif_details_dlg(3)
	If ButtonPressed = add_verif_4_btn Then Call verif_details_dlg(4)
	If ButtonPressed = add_verif_5_btn Then Call verif_details_dlg(5)
	If ButtonPressed = add_verif_6_btn Then Call verif_details_dlg(6)
	If ButtonPressed = add_verif_7_btn Then Call verif_details_dlg(7)
	If ButtonPressed = add_verif_8_btn Then Call verif_details_dlg(8)
	If ButtonPressed = add_verif_9_btn Then Call verif_details_dlg(9)
	If ButtonPressed = add_verif_10_btn Then Call verif_details_dlg(10)
	If ButtonPressed = add_verif_11_btn Then Call verif_details_dlg(11)
	If ButtonPressed = add_verif_12_btn Then Call verif_details_dlg(12)
	If ButtonPressed = add_verif_13_btn Then Call verif_details_dlg(13)
	If ButtonPressed = add_verif_14_btn Then Call verif_details_dlg(14)
	If ButtonPressed = add_verif_15_btn Then Call verif_details_dlg(15)
	If ButtonPressed = add_verif_16_btn Then Call verif_details_dlg(16)
	If ButtonPressed = add_verif_17_btn Then Call verif_details_dlg(17)
	If ButtonPressed = add_verif_18_btn Then Call verif_details_dlg(18)
	If ButtonPressed = add_verif_19_btn Then Call verif_details_dlg(19)
	If ButtonPressed = add_verif_20_btn Then Call verif_details_dlg(20)
	If ButtonPressed = add_verif_21_btn Then Call verif_details_dlg(21)
	If ButtonPressed = add_verif_22_btn Then Call verif_details_dlg(22)
	If ButtonPressed = add_verif_23_btn Then Call verif_details_dlg(23)
	If ButtonPressed = add_verif_24_btn Then Call verif_details_dlg(24)

	If ButtonPressed = open_hsr_manual_transfer_page_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/To_Another_County.aspx"
	If ButtonPressed = add_job_btn Then
		another_job = ""
		count = 0
		for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
			count = count + 1
			If JOBS_ARRAY(jobs_employer_name, each_job) = "" AND JOBS_ARRAY(jobs_employee_name, each_job) = "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) = "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) = "" Then
				another_job = each_job
			End If
		Next
		If another_job = "" Then
			another_job = count
			ReDim Preserve JOBS_ARRAY(jobs_notes, another_job)
		End If
		Call jobs_details_dlg(another_job)
	End If

	for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
		If ButtonPressed = JOBS_ARRAY(jobs_edit_btn, each_job) Then
			Call jobs_details_dlg(each_job)
		End If
	next

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
		If page_display = show_pg_memb_list 			Then ButtonPressed = caf_q_1_6_btn
		If page_display = show_q_1_6 					Then ButtonPressed = caf_q_7_11_btn
		If page_display = show_q_7_11 					Then ButtonPressed = caf_q_12_13_btn
		If page_display = show_q_12_13 					Then ButtonPressed = caf_q_14_15_btn
		If page_display = show_q_14_15 					Then ButtonPressed = caf_q_16_20_btn
		If page_display = show_q_16_20 					Then ButtonPressed = caf_q_21_24_btn
		If page_display = show_q_21_24 					Then ButtonPressed = caf_qual_q_btn
		If page_display = show_qual 					Then ButtonPressed = caf_last_page_btn
		If page_display = show_pg_last 					Then ButtonPressed = finish_interview_btn
		If discrepancies_exist = True Then
			If page_display = show_pg_last 				Then ButtonPressed = discrepancy_questions_btn
			If page_display = discrepancy_questions 	Then ButtonPressed = finish_interview_btn
		End If
		If expedited_determination_needed = True Then
			If expedited_determination_completed = False Then
				If discrepancies_exist = False AND page_display = show_pg_last Then ButtonPressed = expedited_determination_btn
				If page_display = discrepancy_questions 	Then ButtonPressed = expedited_determination_btn
			ElseIf discrepancies_exist = False AND page_display = show_pg_last Then
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
	If ButtonPressed = caf_q_1_6_btn Then
		page_display = show_q_1_6
	End If
	If ButtonPressed = caf_q_7_11_btn Then
		page_display = show_q_7_11
	End If
	If ButtonPressed = caf_q_12_13_btn Then
		page_display = show_q_12_13
	End If
	If ButtonPressed = caf_q_14_15_btn Then
		page_display = show_q_14_15
	End If
	If ButtonPressed = caf_q_16_20_btn Then
		page_display = show_q_16_20
	End If
	If ButtonPressed = caf_q_21_24_btn Then
		page_display = show_q_21_24
	End If
	If ButtonPressed = caf_qual_q_btn Then
		page_display = show_qual
	End If
	If ButtonPressed = caf_last_page_btn Then
		page_display = show_pg_last
	End If
	If ButtonPressed = discrepancy_questions_btn Then
		page_display = discrepancy_questions
	End If
	If ButtonPressed = expedited_determination_btn Then
		' page_display = expedited_determination
		call display_expedited_dialog
	End If

	If ButtonPressed = incomplete_interview_btn Then
		' MsgBox "ARE YOU SURE?"
		confirm_interview_incomplete = MsgBox("You have pressed the button that indicates that the interview was ended but is incomplete." & vbCr & vbCr & "This option is used to end the interview script while clarifying that all interview requirements have not been met." & vbCr & vbCr & "Is this what you want to do?" & vbCr & "(Another dialog will allow you to detail some information about the portion completed.)", vbQuesiton + vbYesNo, "End Interview as Incomplete")
		If confirm_interview_incomplete = vbNo Then
			ButtonPressed = previous_button_pressed
		End If
	End If

	If ButtonPressed = finish_interview_btn or ButtonPressed = incomplete_interview_btn Then leave_loop = TRUE
	If ButtonPressed > 10000 Then
		save_button = ButtonPressed
		If ButtonPressed = page_1_step_1_btn Then call explain_dialog_actions("PAGE 1", "STEP 1")
		If ButtonPressed = page_1_step_2_btn Then call explain_dialog_actions("PAGE 1", "STEP 2")

		ButtonPressed = save_button
	End If

	If ButtonPressed = q_12_all_no_btn Then
		question_12_rsdi_yn = "No"
		question_12_rsdi_amt = ""
		question_12_ssi_yn = "No"
		question_12_ssi_amt = ""
		question_12_va_yn = "No"
		question_12_va_amt = ""
		question_12_ui_yn = "No"
		question_12_ui_amt = ""
		question_12_wc_yn = "No"
		question_12_wc_amt = ""
		question_12_ret_yn = "No"
		question_12_ret_amt = ""
		question_12_trib_yn = "No"
		question_12_trib_amt = ""
		question_12_cs_yn = "No"
		question_12_cs_amt = ""
		question_12_other_yn = "No"
		question_12_other_amt = ""

	End If

	If ButtonPressed = q_14_all_no_btn Then
		question_14_rent_yn = "No"
		question_14_subsidy_yn = "No"
		question_14_mortgage_yn = "No"
		question_14_association_yn = "No"
		question_14_insurance_yn = "No"
		question_14_room_yn = "No"
		question_14_taxes_yn = "No"
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
	                If current_listing = "1"  Then tagline = ": Expedited"        'Adding a specific tagline to the header for the errors
	                If current_listing = "2"  Then tagline = ": CAF ADDR"
	                If current_listing = "3"  Then tagline = ": CAF MEMBs"
	                If current_listing = "4"  Then tagline = ": Q. 1- 6"
	                If current_listing = "5"  Then tagline = ": Q. 7 - 11"
	                If current_listing = "6"  Then tagline = ": Q. 12 - 13"
	                If current_listing = "7"  Then tagline = ": Q. 14 - 15"
					If current_listing = "8"  Then tagline = ": Q. 16 - 20"
					If current_listing = "9"  Then tagline = ": Q. 21 - 24"
					If current_listing = "10" Then tagline = ": CAF QUAL Q"
	                If current_listing = "11" Then tagline = ": CAF Last Page"
					If current_listing = "12" Then tagline = ": Clarifications"
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
				If page_display = show_q_1_6 			Then page_to_review = "4"
				If page_display = show_q_7_11 			Then page_to_review = "5"
				If page_display = show_q_12_13 			Then page_to_review = "6"
				If page_display = show_q_14_15 			Then page_to_review = "7"
				If page_display = show_q_16_20 			Then page_to_review = "8"
				If page_display = show_q_21_24 			Then page_to_review = "9"
				If page_display = show_qual 			Then page_to_review = "10"
				If page_display = show_pg_last			Then page_to_review = "11"
				If page_display = discrepancy_questions Then page_to_review = "12"
				current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing =  trim(current_listing)
				' MsgBox "Page to Review - " & page_to_review & vbCr & "Current Listing - " & current_listing
				If current_listing = page_to_review Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
					If current_listing = "1"  Then tagline = ": Expedited"        'Adding a specific tagline to the header for the errors
					If current_listing = "2"  Then tagline = ": CAF ADDR"
					If current_listing = "3"  Then tagline = ": CAF MEMBs"
					If current_listing = "4"  Then tagline = ": Q. 1- 6"
					If current_listing = "5"  Then tagline = ": Q. 7 - 11"
					If current_listing = "6"  Then tagline = ": Q. 12 - 13"
					If current_listing = "7"  Then tagline = ": Q. 14 - 15"
					If current_listing = "8"  Then tagline = ": Q. 16 - 20"
					If current_listing = "9"  Then tagline = ": Q. 21 - 24"
					If current_listing = "10" Then tagline = ": CAF QUAL Q"
					If current_listing = "11" Then tagline = ": CAF Last Page"
					If current_listing = "12" Then tagline = ": Clarifications"
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
		' If show_err_msg_during_movement = False AND ButtonPressed = finish_interview_btn Then show_msg = True

		' for i = 0 to UBound(HH_MEMB_ARRAY, 2)
		' 	If ButtonPressed = HH_MEMB_ARRAY(button_one, i) Then show_msg = False
		' next
		' If ButtonPressed = update_information_btn Then show_msg = False
		' If ButtonPressed = save_information_btn Then show_msg = False
		' If ButtonPressed = add_person_btn Then show_msg = False
		If page_display = discrepancy_questions Then show_msg = False
		If ButtonPressed = exp_income_guidance_btn Then show_msg = False
		If ButtonPressed = incomplete_interview_btn Then show_msg = False
		If ButtonPressed = verif_button Then show_msg = False
		If ButtonPressed = open_hsr_manual_transfer_page_btn Then show_msg = False
		If ButtonPressed >= 500 AND ButtonPressed < 1200 Then show_msg = False
		If ButtonPressed >= 4000 Then show_msg = False
		' If show_err_msg_during_movement = True AND (ButtonPressed = next_btn OR ButtonPressed = -1) Then show_msg = True
		If error_message = "" Then show_msg = False
		If ButtonPressed = finish_interview_btn Then show_msg = True
		If discrepancies_exist = True AND expedited_determination_needed = False Then
			' MsgBox "1"
			If page_display = discrepancy_questions Then
				' MsgBox "2"
				If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
			End If
		ElseIf expedited_determination_needed = True Then
			' MsgBox "3" & vbCr & "Page Display - ~" & page_display & "~" & vbCr & "LAST Page - ~" & show_pg_last & "~" & vbCr & "exp complete - ~" & expedited_determination_completed& "~"
			If expedited_determination_completed = True AND page_display = show_pg_last Then
				' MsgBox "4"
				If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
			End If
		ElseIf page_display = show_pg_last Then
			' MsgBox "5"
			If ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
		End If
		' MsgBox "Page Display - " & page_display & vbCr & "disc - " & discrepancies_exist & vbCr & "exp det - " & expedited_determination_needed & vbCr & "exp complete - " & expedited_determination_completed & vbCR & "ButtonPressed - " & ButtonPressed & vbCr & "SHOW MSG - " & show_msg
		' MsgBox "Button - " & ButtonPressed & vbCr & "Show? " & show_msg & vbCr & vbCr & "Errors: " & err_msg
		If show_msg = True Then view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")
		If show_msg = False then the_err_msg = ""
        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE AND show_err_msg_during_movement = False Then
            If back_to_dialog = "1"  Then ButtonPressed = caf_page_one_btn         'This calls another function to go to the first dialog that had an error
            If back_to_dialog = "2"  Then ButtonPressed = caf_addr_btn
            If back_to_dialog = "3"  Then ButtonPressed = caf_membs_btn
            If back_to_dialog = "4"  Then ButtonPressed = caf_q_1_6_btn
            If back_to_dialog = "5"  Then ButtonPressed = caf_q_7_11_btn
            If back_to_dialog = "6"  Then ButtonPressed = caf_q_12_13_btn
            If back_to_dialog = "7"  Then ButtonPressed = caf_q_14_15_btn
            If back_to_dialog = "8"  Then ButtonPressed = caf_q_16_20_btn
			If back_to_dialog = "9"  Then ButtonPressed = caf_q_21_24_btn
            If back_to_dialog = "10" Then ButtonPressed = caf_qual_q_btn
            If back_to_dialog = "11" Then ButtonPressed = caf_last_page_btn
            If back_to_dialog = "12" Then ButtonPressed = discrepancy_questions_btn
			If back_to_dialog = "13" Then ButtonPressed = expedited_determination_btn

            Call dialog_movement          'this is where the navigation happens
        End If
    End If
End Function

function complete_MFIP_orientation(CAREGIVER_ARRAY, memb_ref_numb_const, memb_name_const, memb_age_const, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes, family_cash_program)
'DO NOT CHANGE THIS FUNCTION - IT IS DUPLICATED IN AANOTHER SCRIPT AND WE DO NOT WANT TO HAVE TO COMPARE
'*************IMPORTANT - when pulling for FuncLic use the version in DAIL as there are slight changes'

	'first - assess if caregiver meets an exemption
		'- Single parent household employed at least 35 hours per week
		'- 2 Parent household where the 1st parent is employed at least 35 hours per week
		'- 2 Parened household where the 2nd parent is employed at least 20 hours per week and the 1st is employed 35
		'- Pregnant or parenting minor under 20 who is coplying with the educational requirements
		'- Caregiver is not receiving MFIP

	'Identify the caregivers
	'Identify if they are requesting Cash
	'Indicate if this will be DWP or MFIP
	'Identify if the caregiver is a minor
	'List the hours employed for each caregiver
	'
	person_list = "Select One..."+chr(9)+"No Caregiver"
	second_person_list = "Select One..."+chr(9)+"No Second Caregiver"

	For person = 0 to UBound(CAREGIVER_ARRAY, 2)
		person_list = person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
		second_person_list = second_person_list+chr(9)+CAREGIVER_ARRAY(memb_name_const, person)
	Next
	caregiver_one = CAREGIVER_ARRAY(memb_name_const, 0)

	Do
		err_msg = ""
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 551, 150, "Assess for Caregiver MFIP Orientation Requirement"
		  DropListBox 185, 10, 60, 45, "MFIP"+chr(9)+"DWP", family_cash_program
		  EditBox 110, 30, 430, 15, famliy_cash_notes
		  DropListBox 65, 65, 140, 45, person_list, caregiver_one
		  DropListBox 330, 65, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_one_req_cash
		  EditBox 430, 65, 30, 15, caregiver_one_hours_per_week
		  DropListBox 65, 85, 140, 45, second_person_list, caregiver_two
		  DropListBox 330, 85, 45, 45, "Yes"+chr(9)+"No"+chr(9)+"Not Elig", caregiver_two_req_cash
		  EditBox 430, 85, 30, 15, caregiver_two_hours_per_week
		  Text 15, 125, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06   to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
		  ButtonGroup ButtonPressed
			OkButton 490, 125, 50, 15
			PushButton 420, 10, 120, 15, "MFIP Orientation Script Instructions", msg_mfip_orientation_btn
            PushButton 260, 123, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
		  Text 10, 15, 170, 10, "Which Family Cash Program is this Application for?"
		  Text 10, 35, 100, 10, "Notes on Program Selection:"
		  GroupBox 10, 50, 530, 55, "Who are the Caregivers"
		  Text 20, 70, 40, 10, "Caregiver:"
		  Text 215, 70, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 70, 40, 10, "Employed: "
		  Text 465, 70, 50, 10, "hours/week"
		  Text 20, 90, 40, 10, "Caregiver:"
		  Text 215, 90, 115, 10, "Is this caregiver requesting cash?"
		  Text 385, 90, 40, 10, "Employed: "
		  Text 465, 90, 50, 10, "hours/week"
		  Text 15, 110, 100, 10, "Why is this being asked?"
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If caregiver_one = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the First Caregiver or clarify that there is no caregiver"
		If caregiver_two = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate the Second Caregiver or clarify that there is no second caregiver"
		If caregiver_one = caregiver_two Then err_msg = err_msg & vbCr & "* Select two different caregivers"
		If IsNumeric(caregiver_one_hours_per_week) = False AND trim(caregiver_one_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."
		If IsNumeric(caregiver_two_hours_per_week) = False AND trim(caregiver_two_hours_per_week) <> "" Then err_msg = err_msg & vbCr & "* Hours per week should be a number, or left blank."

		If family_cash_program = "DWP" Then err_msg = ""

		If ButtonPressed <> -1 Then err_msg = "LOOP"
		If err_msg <> "" And ButtonPressed = -1 Then MsgBox err_msg

        If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"
		If ButtonPressed = msg_mfip_orientation_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20MFIP%20ORIENTATION.docx"

	Loop until err_msg = ""

	If family_cash_program = "MFIP" Then
		If IsNumeric(caregiver_one_hours_per_week) = True Then caregiver_one_hours_per_week = caregiver_one_hours_per_week * 1
		If trim(caregiver_one_hours_per_week) = "" Then caregiver_one_hours_per_week = 0

		If IsNumeric(caregiver_two_hours_per_week) = True Then caregiver_two_hours_per_week = caregiver_two_hours_per_week * 1
		If trim(caregiver_two_hours_per_week) = "" Then caregiver_two_hours_per_week = 0

		minor_caregiver_on_case = 0

		For person = 0 to UBound(CAREGIVER_ARRAY, 2)
			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_one Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_one_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_one_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_one_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_two <> "No Second Caregiver" AND caregiver_two_req_cash = "Yes" AND caregiver_two_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If

			End If

			If CAREGIVER_ARRAY(memb_name_const, person) = caregiver_two Then
				CAREGIVER_ARRAY(memb_is_caregiver, person) = True
				CAREGIVER_ARRAY(orientation_needed_const, person) = True

				If caregiver_two_req_cash = "Yes" Then CAREGIVER_ARRAY(cash_request_const, person) = True
				If caregiver_two_req_cash <> "Yes" Then
					CAREGIVER_ARRAY(cash_request_const, person) = False
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Caregiver Not on MFIP"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "NO"
				End If
				CAREGIVER_ARRAY(hours_per_week_const, person) = caregiver_two_hours_per_week

				If CAREGIVER_ARRAY(hours_per_week_const, person) > 34 Then
					CAREGIVER_ARRAY(orientation_needed_const, person) = False
					CAREGIVER_ARRAY(orientation_exempt_const, person) = True
					CAREGIVER_ARRAY(exemption_reason_const, person) = "Employed 35+ hours per week"
					CAREGIVER_ARRAY(emps_exemption_code_const, person) = "20"
				ElseIf CAREGIVER_ARRAY(hours_per_week_const, person) > 19 Then
					If caregiver_one <> "No Second Caregiver" AND caregiver_one_req_cash = "Yes" AND caregiver_one_hours_per_week > 34 Then
						CAREGIVER_ARRAY(orientation_needed_const, person) = False
						CAREGIVER_ARRAY(orientation_exempt_const, person) = True
						CAREGIVER_ARRAY(exemption_reason_const, person) = "2nd Caregiver Employed 20+ hours per week"
						CAREGIVER_ARRAY(emps_exemption_code_const, person) = "21"
					End If
				End If
				If CAREGIVER_ARRAY(memb_age_const, person) < 20 Then
					minor_caregiver_on_case = minor_caregiver_on_case + 1
					CAREGIVER_ARRAY(exempt_from_ed_const, person) = "No"
					CAREGIVER_ARRAY(comply_with_ed_const, person) = "Yes"
				End If
			End If



		Next

		'IF A MINOR IS FOUND
		If minor_caregiver_on_case > 0 Then
			Do
				err_msg = ""
				dlg_len = 210
				If minor_caregiver_on_case = 2 Then dlg_len = 290

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 551, dlg_len, "Assess for Caregiver MFIP Orientation Requirement"
				  Text 10, 15, 200, 10, "Which Family Cash Program is this Application for? " & family_cash_program
				  Text 10, 25, 500, 20, "Notes on Program Selection: " & famliy_cash_notes
				  GroupBox 10, 50, 530, 40, "Who are the Caregivers"
				  Text 20, 60, 190, 10, "Caregiver: " & caregiver_one
				  Text 215, 60, 165, 10, "Is this caregiver requesting cash? " & caregiver_one_req_cash
				  Text 385, 60, 90, 10, "Employed: " & caregiver_one_hours_per_week
				  Text 465, 60, 50, 10, "hours/week"
				  Text 20, 75, 190, 10, "Caregiver: " & caregiver_two
				  Text 215, 75, 165, 10, "Is this caregiver requesting cash? " & caregiver_two_req_cash
				  Text 385, 75, 90, 10, "Employed: " & caregiver_two_hours_per_week
				  Text 465, 75, 50, 10, "hours/week"
				  y_pos = 30
				  For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
					  If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
						  y_pos = y_pos + 70
						  GroupBox 10, y_pos, 530, 65, CAREGIVER_ARRAY(memb_name_const, caregiver)
						  Text 20, y_pos + 10, 270, 10, "This caregiver appears to be a minor by MFIP program rules (under 20 years old)."
						  Text 20, y_pos + 30, 195, 10, "Is this caregiver exempt from the Educational Requirement?"
						  DropListBox 230, y_pos + 25, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(exempt_from_ed_const, caregiver)
						  Text 20, y_pos + 50, 205, 10, "Is this caregiver complying with the Educational Requirement?"
						  DropListBox 230, y_pos + 45, 40, 45, "No"+chr(9)+"Yes", CAREGIVER_ARRAY(comply_with_ed_const, caregiver)
					  End If
				  Next
				  Text 15, y_pos + 90, 450, 20, "These questions will identify if these caregivers need an MFIP orientation. See CM 05.12.12.06 to see the reasons that a caregiver would not need an MFIP Orientation. The script will use this information to determine if the MFIP Orientation Functionality should be run."
				  ButtonGroup ButtonPressed
					OkButton 490, y_pos + 90, 50, 15
					PushButton 485, y_pos + 45, 50, 15, "CM 28.12", cm_28_12_btn
					PushButton 260, y_pos + 87, 55, 10, "CM05.12.12.06", cm_05_12_12_06_btn
				  Text 355, y_pos + 45, 125, 20, "See details about the educational requirement in the Combined Manual "
				  Text 15, y_pos + 75, 100, 10, "Why is this being asked?"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If err_msg <> "" Then MsgBox err_msg

			Loop until err_msg = ""

			For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)
				If CAREGIVER_ARRAY(memb_is_caregiver, caregiver) = True and CAREGIVER_ARRAY(memb_age_const, caregiver) < 20 Then
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = True
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "No" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
					If CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = "Yes" Then CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True

					If CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False and CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = True Then
						CAREGIVER_ARRAY(orientation_needed_const, caregiver) = False
						CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True
						CAREGIVER_ARRAY(exemption_reason_const, caregiver) = "Minor Caregiver meeting Educational Requirements"
						CAREGIVER_ARRAY(emps_exemption_code_const, caregiver) = "22"
					End If
				Else
					CAREGIVER_ARRAY(exempt_from_ed_const, caregiver) = False
					CAREGIVER_ARRAY(comply_with_ed_const, caregiver) = False
				End If
			Next

		End If

		const mf_step_rights_resp 	= 1
		const mf_step_time_limits	= 2
		const mf_step_extension		= 3
		const mf_step_dv			= 4
		const mf_step_expectations	= 5
		const mf_step_esp			= 6
		const mf_step_compliance	= 7
		const mf_step_ep			= 8
		const mf_step_ccap			= 9
		const mf_step_incentives	= 10
		const mf_step_hc			= 11
		const mf_completion			= 12

		' mf_step_rights_resp_viewed = False
		' mf_step_time_limits_viewed = False
		' mf_step_extension_viewed = False
		' mf_step_dv_viewed = False
		' mf_step_expectations_viewed = False
		' mf_step_esp_viewed = False
		' mf_step_compliance_viewed = False
		' mf_step_ep_viewed = False
		' mf_step_ccap_viewed = False
		' mf_step_incentives_viewed = False
		' mf_step_hc_viewed = False
		' mf_completion_viewed = False
		'
		' orientation_script_document_viewed = False
		'
		'FIRST - Participant Responsibilities and Rights'
		'SECOND - MFIP Time Limits'
		'THIRD - MFIp Extension Eligibility'
		'FOURTH - Family Violence'
		'FIFTH - Expectations'
		'SIXTH - Choosing ESP'
		'SEVENTH - Assignment and Compliance'
		'EIGHTH - Developing an EP'
		'NINTH - CCAP'
		'TENTH - Incentives'
		'ELEVENTH - Health Care'

		' all_mfip_orientation_info_viewed = False
		For caregiver = 0 to UBound(CAREGIVER_ARRAY, 2)

			If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then
                Call Navigate_to_MAXIS_screen("STAT", "EMPS")
    			If CAREGIVER_ARRAY(memb_ref_numb_const, caregiver) <> "" Then
    				EMWriteScreen CAREGIVER_ARRAY(memb_ref_numb_const, caregiver), 20, 76
    				transmit
    			End If

                MFIP_orientation_step = mf_step_rights_resp

				mf_step_rights_resp_viewed = False
				mf_step_time_limits_viewed = False
				mf_step_extension_viewed = False
				mf_step_dv_viewed = False
				mf_step_expectations_viewed = False
				mf_step_esp_viewed = False
				mf_step_compliance_viewed = False
				mf_step_ep_viewed = False
				mf_step_ccap_viewed = False
				mf_step_incentives_viewed = False
				mf_step_hc_viewed = False
				mf_completion_viewed = False

				orientation_script_document_viewed = False

				all_mfip_orientation_info_viewed = False

				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 551, 385, "MFIP Orientation"
					  ' GroupBox 10, 10, 450, 45, "Group1"
					  ButtonGroup ButtonPressed
					  	If MFIP_orientation_step <> mf_completion Then PushButton 495, 365, 50, 15, "NEXT", next_btn

						'FIRST - Participant Responsibilities and Rights'
						If MFIP_orientation_step = mf_step_rights_resp Then
						  Text 10, 10, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  GroupBox 10, 30, 450, 130, "Participant Responsibilities and Rights"
						  Text 20, 45, 370, 10, "As a program participant you have responsibilities and rights that were discussed during your intake interview."
						  Text 20, 60, 430, 10, "Please keep a copy of the Client Responsibilities and Rights (DHS-4163) for your reference. Let us know if you have any questions."
						  Text 20, 80, 335, 10, "Please remember it's important to report ANY changes that could affect your eligibility within 10 days."
						  GroupBox 10, 75, 450, 15, ""
						  Text 20, 100, 420, 20, "If your income decreases by at least 50% contact your financial worker right away!  You may be eligible for a significant change meaning a recalculation of your income which may result in an increase of your cash and/or food benefits."
						  Text 20, 125, 335, 20, "If you do not meet program eligibility such as cash assistance, your financial worker will assess other program eligibility such as SNAP."
						  PushButton 385, 160, 75, 15, "DHS - 4163", open_dhs_4163_btn

						  mf_step_rights_resp_viewed = True
						  'ADD BUTTON DHS 4163'
						End If

						'SECOND - MFIP Time Limits'
						If MFIP_orientation_step = mf_step_time_limits Then
						  GroupBox 10, 10, 450, 160, "MFIP Time-Limits"
						  Text 20, 25, 430, 30, "The MFIP program is available to you for up to 60 months in your lifetime.  If you have used cash assistance in another state those months must be reported and may count toward your lifetime limit. There are some instances the months you use may be exempt, meaning the months do not count towards the 60-month lifetime limit."
						  Text 20, 55, 55, 10, "These Include:"
						  Text 30, 70, 125, 10, "1. Months you are over 60 years old"
						  Text 30, 80, 310, 10, "2. Months you are living on a reservation where at least 50% of the adults were not employed"
						  Text 30, 90, 360, 10, "3. Months when you are a victim of family violence AND have an approved family violence waiver plan"
						  Text 30, 100, 335, 10, "4. Months you don't receive the cash portion of MFIP (*talk to your financial worker for more details)"
						  Text 30, 110, 350, 10, "5. Months you are a parent under 18 years of age and complying with your school or social service plan"
						  Text 30, 120, 395, 10, "6. Months you are 18 or 19 years old and do not have a high school diploma/GED AND complying with a school plan"
						  Text 40, 135, 355, 25, "Note: If you are eligible for an exemption but you are not complying with program requirements and do not meet a good cause reason, those months will count toward the lifetime limit. If you have questions about possible good cause reasons, talk to a worker."
							  Text 10, 175, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_time_limits_viewed = True
						End If

						'THIRD - MFIp Extension Eligibility'
						If MFIP_orientation_step = mf_step_extension Then
						  GroupBox 10, 10, 450, 295, "MFIP Extension Eligibility"
						  Text 20, 30, 165, 10, "You may be eligible for an MFIP Extension if:"
						  Text 30, 45, 375, 10, "- You are a single or two parent household working the required number of hours that meet extension eligibility"
						  Text 30, 55, 365, 10, "- Your health care provider states you are only able to work 20 hours per week due to an illness or disability"
						  Text 20, 75, 380, 20, "A qualified professional verifies you have one or more of the conditions below that severely limits your ability to obtain or maintain suitable employment for 20 or more hours per week:"
						  Text 30, 100, 165, 10, "- Developmentally Disabled or Mentally Ill"
						  Text 30, 110, 95, 10, "- Learning Disability"
						  Text 30, 120, 60, 10, "- IQ Below 80"
						  Text 30, 130, 260, 10, "- You are ill/injured or incapacitated that's expected to last more than 30 days"
						  Text 20, 145, 125, 10, "A qualified professional verifies:"
						  Text 35, 160, 280, 15, "You are needed in the home to provide care for a family member or foster child in the household that is expected to continue for more than 30 days "
						  Text 35, 185, 285, 35, "A child or adult in the home meets the Special Medical Criteria for home care services or a home and community-based waiver services program, severe emotional disturbance (SED diagnosed child) or serious and persistent mental illness (SPMI diagnosed adult)"
						  Text 35, 225, 275, 20, "You have significant barriers to employment and determined Unemployable by a vocational specialist or other qualified professional designated by the county"
						  Text 35, 250, 165, 10, "You are a victim of family violence"
						  Text 20, 265, 415, 30, "If you believe you meet any of the criteria's above it's important to discuss with your financial worker AND your employment counselor. You may qualify for a modified employment plan prior to reaching your 60-month as well as receive an extension of your cash benefits."
							  Text 10, 310, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_extension_viewed = True
						End If

						'FOURTH - Family Violence'
						If MFIP_orientation_step = mf_step_dv Then
						  GroupBox 10, 10, 450, 75, "Family Violence Resources/Supports"
						  Text 20, 30, 390, 10, "Your financial worker discussed and provided information regarding resources if you are a victim of family violence."
						  Text 20, 45, 375, 35, " Please review that brochure if you need assistance with shelter and/or supports Domestic Violence Information (DHS 3477) and Family Violence Referral (DHS 3323). If you are a victim of domestic violence, you may choose to work with your assigned Employment Counselor to determine if you are eligible for a Family Violence Waiver to allow your family time and flexibility to focus on safety issues."
							  Text 10, 90, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 90, 75, 15, "DHS - 3477", open_dhs_3477_btn
						  PushButton 385, 105, 75, 15, "DHS - 3323", open_dhs_3323_btn

						  mf_step_dv_viewed = True
						  'ADD BUTTON DHS 3477
						  'ADD BUTTON DHS 3323
						End If

						'FIFTH - Expectations'
						If MFIP_orientation_step = mf_step_expectations Then
						  GroupBox 10, 10, 450, 110, "Expectations of Participants Approved for the MFIP Program"
						  Text 20, 30, 360, 20, "MFIP services focus on putting you on the most direct path to employment and other related steps that will support long-term economic stability."
						  Text 20, 55, 375, 20, "While you are expected to work, look for work, or participate in activities to prepare for work, the steps toward economic stability look different for all families and participants."
						  Text 20, 80, 405, 20, "Employment Services have a variety of tools to address the unique needs of each family. You will hear more about these tools and resources during your Employment Services Overview."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_expectations_viewed = True
						End If

						'SIXTH - Choosing ESP'
						If MFIP_orientation_step = mf_step_esp Then
						  GroupBox 10, 10, 450, 155, "Choosing an MFIP Employment Service Provider (ESP)"
						  Text 25, 25, 345, 10, "As part of the MFIP program you are required to work with an MFIP Employment Service Provider (ESP)."
						  Text 25, 40, 410, 20, "There's a variety of providers available to help support your employment goals. On the MFIP ESP Choice Sheet, choose the top three providers you'd like to work with listing your top three choices with 1 being the provider you most want to work with."
						  Text 25, 65, 330, 10, "We will do our best to refer you to one of your top three choices depending on available openings."
						  Text 25, 85, 195, 10, "There are a few exceptions in choosing your provider:"
						  Text 40, 100, 345, 10, "If you have worked with an MFIP ESP in the past ninety (90) days, you may be referred to that provider."
						  Text 40, 115, 350, 20, "If you are under 18 and do not have a HS diploma/GED, you will be referred to Minnesota Visiting Nurse Association to discuss your education and employment options"
						  Text 40, 140, 345, 20, "If you have used 60 months or more of your TANF time limit and granted an extension under a specific category you will be referred to an agency that specializes in that type of extension."
							  Text 10, 170, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  ' PushButton 385, 170, 75, 15, "Choice Sheet", open_choice_sheet_btn

						  mf_step_esp_viewed = True
						  'ADD BUTTON - CHOICE SHEET ???'
						End If

						'SEVENTH - Assignment and Compliance'
						If MFIP_orientation_step = mf_step_compliance Then
						  GroupBox 10, 10, 450, 110, "Assignment and Compliance with MFIP Employment Services"
						  Text 25, 30, 295, 10, "Once you are approved for MFIP you will be referred to an Employment Service Provider."
						  Text 25, 45, 375, 20, "In Hennepin County, many of the Employment Services Providers are community based nonprofit organizations who partner with Hennepin County to deliver services."
						  Text 25, 70, 410, 20, "The provider will send you a notice to attend an MFIP Employment Service Overview. You are required to attend the overview and work with your assigned employment service counselor."
						  Text 25, 95, 400, 20, "If you choose not to comply with program requirements, your case may be sanctioned resulting in a reduction of your cash and/or food benefits."
							  Text 10, 125, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_compliance_viewed = True
						End If

						'EIGHTH - Developing an EP'
						If MFIP_orientation_step = mf_step_ep Then
						  GroupBox 10, 10, 450, 270, "Developing an Employment Plan (EP) with your MFIP Employment Counselor"
						  Text 20, 25, 430, 25, "Program participants will work with their assigned Employment Counselor to develop an Employment Plan. Your Employment Plan will be based on your goals and will include activities that are intended to lead to employment and financial stability. On the path to stable employment, many different types of activities are available."
						  Text 20, 55, 140, 10, "Some of the allowable activities include:"
						  Text 30, 70, 260, 10, "- Job search (including participation in job clubs, workshops, and hiring events)"
						  Text 30, 80, 260, 10, "- Employment"
						  Text 30, 90, 260, 10, "- Self-employment"
						  Text 30, 100, 260, 10, "- Community work experience and/or volunteer work"
						  Text 30, 110, 260, 10, "- On the job training"
						  Text 30, 120, 260, 10, "- English Language Learning (ELL and ESL) or Functional Work Literacy (FWL)"
						  Text 30, 130, 260, 10, "- Adult Basic Education, GED preparation and Adult High School Diploma"
						  Text 30, 140, 260, 10, "- Job skills training directly related to employment"
						  Text 30, 150, 260, 10, "- Post-Secondary Training and Education"
						  Text 30, 160, 415, 10, "- Other activities that are critical to your family's success in reaching your employment goals such as chemical dependency"
						  Text 35, 170, 260, 10, "treatment, mental health services, social services, and parenting education."
						  Text 20, 190, 430, 25, "You are required to follow through with the activities in your employment plan. If you are unable to complete the activities, contact your Employment Counselor right away to determine if your plan need to be updated. Good communication with your employment counselor can help prevent reduction in your grant (sanctions)."
						  Text 20, 220, 425, 30, "Your Employment Counselor may conduct assessments with you to support you in selecting an education and training path that creates opportunities for long term economic stability. If you have more questions about education and training options, you can also see the Education and Training Brochure (DHS 3366)."
						  Text 20, 255, 420, 20, "Work study programs under the higher education systems may also be available.  Your assigned employment counselor will discuss this opportunity in more detail."
							  Text 10, 285, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 285, 75, 15, "DHS - 3366", open_dhs_3366_btn

						  mf_step_ep_viewed = True
						  'ADD BUTTON DHS 3366'
						ENd If

						'NINTH - CCAP'
						If MFIP_orientation_step = mf_step_ccap Then
						  GroupBox 10, 10, 450, 150, "Availability of Childcare Assistance"
						  Text 25, 25, 425, 20, "There are several Childcare Assistance programs (CCAP) available to support your participation in employment, pre-employment activities, training, and/or educational programs"
						  Text 40, 45, 120, 10, "- MFIP/DWP Childcare assistance"
						  Text 40, 55, 120, 10, "- Transition Year Childcare assistance"
						  Text 60, 65, 385, 25, "Many families continue to be eligible for childcare assistance when their MFIP case closes.  It's highly recommended that you speak to your assigned childcare worker to discuss eligibility details specific to your continued needs for assistance when MFIP closes"
						  Text 40, 90, 170, 10, "- Transition Year Extension Childcare assistance"
						  Text 40, 100, 170, 10, "- Basic sliding fee Childcare assistance"
						  Text 60, 110, 215, 10, "If funds are not available, you may be put on a waiting list"
						  Text 25, 125, 430, 10, "Contact your assigned Employment Counselor or Childcare Assistance Worker to discuss eligibility requirements in more detail."
						  Text 25, 140, 395, 10, "If you need help locating childcare provider options, here's a great resource to contact Think Small or (651-641-0332)"
						  GroupBox 10, 165, 450, 65, "Who to Contact about Childcare Assistance?"
						  Text 25, 180, 420, 20, "If you are receiving MFIP your assigned Employment Counselor will work with you to determine how many childcare hours need to be approved based on the activities in your Employment Plan"
						  Text 25, 205, 420, 20, "If you are receiving MFIP but have not been assigned to an Employment Counselor or if your MFIP has closed contact the childcare assistance line directly at 612-348-5937"
						  GroupBox 10, 235, 450, 115, "Program Compliance and Unavailability of Childcare Assistance"
						  Text 25, 250, 425, 20, "The county may NOT impose a sanction for failure to comply with program requirements if you have good cause because of the unavailability of childcare. The inability to obtain childcare does not exempt or extend your TANF time limit."
						  Text 25, 275, 105, 10, "Some good cause reasons are:"
						  Text 35, 285, 135, 10, "- Unavailability of appropriate childcare"
						  Text 35, 295, 135, 10, "- Unreasonable distance to childcare provider"
						  Text 35, 305, 235, 10, "- Provider does not meet health and safety standards for the child(ren)"
						  Text 35, 315, 275, 10, "- The provider charges an excess amount above the maximum the county can pay"
						  Text 25, 330, 335, 10, "Your Childcare Worker or Employment Counselor can discuss good cause reasons in more detail"
							  Text 10, 365, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)

						  mf_step_ccap_viewed = True
						End If

						'TENTH - Incentives'
						If MFIP_orientation_step = mf_step_incentives Then
						  GroupBox 10, 10, 450, 150, "Incentives and Tax Credits"
						  Text 25, 30, 230, 10, "The MFIP program is designed to benefit you when you are working."
						  Text 25, 45, 420, 35, "For example, your financial worker will not budget all your earned income when they calculate the amount of cash and food benefits you are eligible for. When determining your benefit amount, they will not count the first $65 of income you earn AND beyond that, they will only count half of your remaining gross earned income. Here is a link to explain how this works: Bulletin 21-11-01 - DHS Reissues 'Work Will Always Pay ... With MFIP'"
						  Text 25, 85, 425, 10, "If you are working, when you file your taxes apply for the Earned Income Credit and the Minnesota Working Family Credit."
						  Text 25, 100, 225, 10, "Getting a tax refund will NOT affect your eligibility for MFIP."
						  Text 25, 115, 425, 35, "Have your taxes done for FREE! For a list of free tax preparation sites call the Minnesota Department of Revenue at 651-296-3781 or 1-800-652-9094. Neighborhood Volunteer Income Tax Assistance (VITA) sites are available throughout the state. They are open from February 1 through April 15. Some sites are open year around to help you file back taxes. Search for free tax preparation sites at Department of Revenue."
							  Text 10, 165, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 165, 75, 15, "DHS Bulletin 21-11-01", open_dhs_bulletin_21_11_01_btn

						  mf_step_incentives_viewed = True
						  'ADD BUTTON BULLETIN 21-11-01'
						End If

						'ELEVENTH - Health Care'
						If MFIP_orientation_step = mf_step_hc Then
						  GroupBox 10, 10, 450, 90, "Health Care"
						  Text 25, 30, 230, 10, "You may qualify for Minnesota Health Care programs."
						  Text 25, 45, 410, 20, "You can apply for health care online at www.mnsure.org (for assistance completing an online application call 1-855-366-7873) or we can mail you a paper application (DHS 6696)."
						  Text 25, 70, 425, 20, "For help with age-appropriate preventive health services check out the Child and Teen Checkup program at: http://edocs.dhs.state.mn.us/lfserver/public/DHS-1826-ENG"
							  Text 10, 105, 145, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
						  PushButton 385, 105, 75, 15, "DHS - 1826", open_dhs_1826_btn

						  mf_step_hc_viewed = True
						  'ADD BUTTON DHS 1826'
						End If

						If MFIP_orientation_step = mf_completion Then
						  GroupBox 10, 10, 450, 140, "Document MFIP Orientation Completion"
						  Text 20, 30, 135, 10, "For " & CAREGIVER_ARRAY(memb_name_const, caregiver) & ":"
						  Text 25, 50, 215, 10, "Did you verbally review all information in the MFIP Oreientation?"
						  DropListBox 240, 45, 210, 45, "Select One..."+chr(9)+"Yes - all information has been reviewed"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(orientation_done_const, caregiver)
						  Text 25, 65, 240, 10, "Notes from any questions/conversation during the MFIP Orientation:"
						  EditBox 25, 75, 425, 15, CAREGIVER_ARRAY(orientation_notes, caregiver)
						  Text 25, 105, 125, 10, "IF COMPLETE - OPEN ECF NOW"
						  Text 35, 120, 220, 10, "Complete the ESP Choice Sheet (D387) with the resident now."
						  Text 35, 135, 175, 10, "Confirm Choice Sheet Completed and saved to ECF:"
						  DropListBox 205, 130, 140, 45, "Select One..."+chr(9)+"Yes - Choice Sheet Saved to ECF"+chr(9)+"No - could not complete", CAREGIVER_ARRAY(choice_form_done_const, caregiver)
						  Text 210, 155, 250, 25, "MFIP Orientation is now complete for this resident. If this case has a second caregiver that requires the MFIP Orientation, this dialog will reappear for the next caregiver as this is a person based process."
						  PushButton 385, 180, 75, 15, "HSR Manual", open_hsr_manual_btn
						  PushButton 385, 195, 75, 15, "CM 05.12.12.06", cm_05_12_12_06_btn

						  mf_completion_viewed = True
						End If

						Text 470, 5, 80, 10, "MFIP Orientation Topics"
						Text 10, 360, 190, 20, "The entire MFIP Orientation to Financial Serviews Script can be viewed on Sharepoint - Open Word Document here:"

						If MFIP_orientation_step = mf_step_rights_resp Then 	Text 500, 18, 55, 10, "Rights / Resp"
						If MFIP_orientation_step = mf_step_time_limits Then 	Text 504, 33, 55, 10, "Time Limits"
						If MFIP_orientation_step = mf_step_extension Then 	Text 509, 48, 55, 10, "Extention"
						If MFIP_orientation_step = mf_step_dv Then 			Text 497, 63, 55, 10, "Family Violence"
						If MFIP_orientation_step = mf_step_expectations Then 	Text 503, 78, 55, 10, "Expectations"
						If MFIP_orientation_step = mf_step_esp Then 			Text 508, 93, 55, 10, "MFIP ESP"
						If MFIP_orientation_step = mf_step_compliance Then 	Text 497, 108, 55, 10, "ES Compliance"
						If MFIP_orientation_step = mf_step_ep Then 			Text 505, 123, 55, 10, "Emplmt Plan"
						If MFIP_orientation_step = mf_step_ccap Then 			Text 512, 138, 55, 10, "CCAP"
						If MFIP_orientation_step = mf_step_incentives Then 	Text 506, 153, 55, 10, "Incentives"
						If MFIP_orientation_step = mf_step_hc Then 			Text 505, 168, 55, 10, "Health Care"
						If MFIP_orientation_step = mf_completion Then 		Text 502, 188, 55, 10, "Confirmation"


					    If MFIP_orientation_step = mf_completion Then PushButton 495, 365, 50, 15, "DONE", done_btn


					    If MFIP_orientation_step <> mf_step_rights_resp Then 	PushButton 495, 15, 55, 15, "Rights / Resp", button_one
					    If MFIP_orientation_step <> mf_step_time_limits Then 	PushButton 495, 30, 55, 15, "Time Limits", button_two
						If MFIP_orientation_step <> mf_step_extension Then 		PushButton 495, 45, 55, 15, "Extention", button_three
					    If MFIP_orientation_step <> mf_step_dv Then 			PushButton 495, 60, 55, 15, "Family Violence", button_four
					    If MFIP_orientation_step <> mf_step_expectations Then 	PushButton 495, 75, 55, 15, "Expectations", button_five
					    If MFIP_orientation_step <> mf_step_esp Then 			PushButton 495, 90, 55, 15, "MFIP ESP", button_six
					    If MFIP_orientation_step <> mf_step_compliance Then 	PushButton 495, 105, 55, 15, "ES Compliance", button_seven
					    If MFIP_orientation_step <> mf_step_ep Then 			PushButton 495, 120, 55, 15, "Emplmt Plan", button_eight
					    If MFIP_orientation_step <> mf_step_ccap Then 			PushButton 495, 135, 55, 15, "CCAP", button_nine
					    If MFIP_orientation_step <> mf_step_incentives Then 	PushButton 495, 150, 55, 15, "Incentives", button_ten
					    If MFIP_orientation_step <> mf_step_hc Then 			PushButton 495, 165, 55, 15, "Health Care", button_eleven
					    If MFIP_orientation_step <> mf_completion Then 			PushButton 495, 185, 55, 15, "Confirmation", button_twelve

					    ' PushButton 495, 195, 55, 15, "Button 2", Button13
					    ' PushButton 495, 210, 55, 15, "Button 2", Button14
					    ' PushButton 495, 225, 55, 15, "Button 2", Button15
					    ' PushButton 495, 240, 55, 15, "Button 2", Button16
						PushButton 205, 360, 135, 15, "MFIP Oriendation Document", mfip_orientation_word_doc_btn
						' OkButton 495, 365, 50, 15

					EndDialog

					dialog Dialog1
					cancel_confirmation

					err_msg = ""

					If ButtonPressed = next_btn Then MFIP_orientation_step = MFIP_orientation_step + 1
					If ButtonPressed = button_one Then MFIP_orientation_step = mf_step_rights_resp
					If ButtonPressed = button_two Then MFIP_orientation_step = mf_step_time_limits
					If ButtonPressed = button_three Then MFIP_orientation_step = mf_step_extension
					If ButtonPressed = button_four Then MFIP_orientation_step = mf_step_dv
					If ButtonPressed = button_five Then MFIP_orientation_step = mf_step_expectations
					If ButtonPressed = button_six Then MFIP_orientation_step = mf_step_esp
					If ButtonPressed = button_seven Then MFIP_orientation_step = mf_step_compliance
					If ButtonPressed = button_eight Then MFIP_orientation_step = mf_step_ep
					If ButtonPressed = button_nine Then MFIP_orientation_step = mf_step_ccap
					If ButtonPressed = button_ten Then MFIP_orientation_step = mf_step_incentives
					If ButtonPressed = button_eleven Then MFIP_orientation_step = mf_step_hc
					If ButtonPressed = button_twelve Then MFIP_orientation_step = mf_completion


					If ButtonPressed = mfip_orientation_word_doc_btn Then
						run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-es-manual/_layouts/15/Doc.aspx?sourcedoc=%7BCB2C8281-95F1-45EE-84D8-B2DF617AA62C%7D&file=MFIP%20Orientation%20to%20Financial%20Services.docx"
						MFIP_orientation_step = mf_completion
						orientation_script_document_viewed = True
					End If
					If ButtonPressed = open_dhs_4163_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
					If ButtonPressed = open_dhs_3477_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
					If ButtonPressed = open_dhs_3323_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3323-ENG"
					If ButtonPressed = open_dhs_3366_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3366-ENG"
					If ButtonPressed = open_dhs_bulletin_21_11_01_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_FILE&RevisionSelectionMethod=LatestReleased&Rendition=Primary&allowInterrupt=1&noSaveAs=1&dDocName=dhs-328254"
					If ButtonPressed = open_dhs_1826_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-1826-ENG"

					If ButtonPressed = open_hsr_manual_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/MFIP_Orientation.aspx"
					If ButtonPressed = cm_05_12_12_06_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_0005121206"
					' If ButtonPressed = cm_28_12_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=CM_002812"




					If mf_step_rights_resp_viewed = True and mf_step_time_limits_viewed = True and mf_step_extension_viewed = True and mf_step_dv_viewed = True and mf_step_expectations_viewed = True and mf_step_esp_viewed = True and mf_step_compliance_viewed = True and mf_step_ep_viewed = True and mf_step_ccap_viewed = True and mf_step_incentives_viewed = True and mf_step_hc_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True
					If orientation_script_document_viewed = True and mf_completion_viewed = True Then all_mfip_orientation_info_viewed = True


					' MsgBox "DONE? - " & CAREGIVER_ARRAY(orientation_done_const, caregiver) & vbCr & "CHOICE SHEET? - " & CAREGIVER_ARRAY(choice_form_done_const, caregiver)
					If all_mfip_orientation_info_viewed = False and CAREGIVER_ARRAY(orientation_done_const, caregiver) = "No - could not complete" Then err_msg = err_msg & vbCr & "* You must review the entire MFIP Orientation before continuing."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP Orientation has been completed."
					If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" and CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Select One..." Then err_msg = err_msg & vbCr & "* Indicate if the MFIP ESP Choice Sheet has been completed in ECF."

					If ButtonPressed = done_btn and err_msg <> "" Then MsgBox err_msg
					' If ButtonPressed = done_btn Then MsgBox err_msg
					If ButtonPressed <> done_btn Then err_msg = "HOLD"

				Loop Until all_mfip_orientation_info_viewed = True and err_msg = ""
			End If
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "Yes - all information has been reviewed" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = True
			If CAREGIVER_ARRAY(orientation_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(orientation_done_const, caregiver) = False
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "Yes - Choice Sheet Saved to ECF" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True
			If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = "No - could not complete" Then CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False

			'HERE WE HAVE A DIALOG TO GO TO EMPS AND GIVE INSTRUCTION ON HOW TO COMPLETE IT
			If (CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True and CAREGIVER_ARRAY(orientation_done_const, caregiver) = True) or CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 281, 185, "Update EMPS Panel"
				  ButtonGroup ButtonPressed
				    PushButton 125, 135, 145, 15, "The EMPS Panel Update is Complete", emps_update_complete_btn
				  Text 15, 10, 125, 10, "Caregiver: " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  If CAREGIVER_ARRAY(orientation_needed_const, caregiver) = True Then Text 35, 20, 205, 10, "NEEDS an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 35, 20, 205, 10, "Is Exempt from having an MFIP Orientation"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is Completed"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 15, 35, 255, 10, "The MFIP Orientation to Financial Services is NOT Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is Completed"
				  If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = False Then Text 15, 45, 255, 10, "The ESP Choice Sheet in ECF is NOT Completed"

				  Text 15, 65, 260, 10, "This person has met the requirement for the MFIP Orientation."
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 20, 75, 260, 10, "Exemption Reason: " & CAREGIVER_ARRAY(exemption_reason_const, caregiver)
				  GroupBox 15, 90, 255, 65, "Update EMPS Panel Now"
				  Text 25, 105, 210, 10, "Update panel: EMPS for " & CAREGIVER_ARRAY(memb_name_const, caregiver)
				  Text 30, 115, 45, 10, "Fin Orient Dt: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 115, 40, 10, date
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 115, 40, 10, "__ __ __"
				  Text 45, 125, 35, 10, "Attended: "
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then Text 85, 125, 20, 10, "Y"
				  If CAREGIVER_ARRAY(orientation_done_const, caregiver) = False Then Text 85, 125, 20, 10, "N"
				  Text 30, 135, 45, 10, "Good Cause:"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = False Then Text 85, 135, 20, 10, "__"
				  If CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then Text 85, 135, 20, 10, CAREGIVER_ARRAY(emps_exemption_code_const, caregiver)
				EndDialog

				dialog Dialog1

				Call start_a_blank_CASE_NOTE

				If CAREGIVER_ARRAY(orientation_done_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE("MFIP Orientation completed with " & CAREGIVER_ARRAY(memb_name_const, caregiver))
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Completed on", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Orientation Notes", CAREGIVER_ARRAY(orientation_notes, caregiver))
					If CAREGIVER_ARRAY(choice_form_done_const, caregiver) = True Then Call write_variable_in_CASE_NOTE("* ESP Choice Sheet: Completed in Case File ")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " did not meet an exemption from completing an MFIP Orientation")
					Call write_variable_in_CASE_NOTE("---")
                    Call write_bullet_and_variable_in_CASE_NOTE("Notes on Program Selection", famliy_cash_notes)
                    Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				ElseIf CAREGIVER_ARRAY(orientation_exempt_const, caregiver) = True Then

					Call write_variable_in_CASE_NOTE(CAREGIVER_ARRAY(memb_name_const, caregiver) & " is Exempt from MFIP Orientation")
					Call write_bullet_and_variable_in_CASE_NOTE("Assessment Completed", date)
					Call write_bullet_and_variable_in_CASE_NOTE("Exemption Reason", CAREGIVER_ARRAY(exemption_reason_const, caregiver))
					Call write_variable_in_CASE_NOTE("---")
                    Call write_bullet_and_variable_in_CASE_NOTE("Notes on Program Selection", famliy_cash_notes)
                    Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature)

				End If
				PF3

                call back_to_SELF

			End If
			' MsgBox CAREGIVER_ARRAY(memb_name_const, caregiver) & " - DONE"
		Next
	End If
	' MsgBox "STOP HERE"
end function

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
	hsr_snap_applications_btn		= 1100
	ryb_exp_identity_btn			= 1200
	ryb_exp_timeliness_btn			= 1300
	sir_exp_flowchart_btn			= 1400
	cm_04_04_btn					= 1500
	cm_04_06_btn					= 1600
	ht_id_in_solq_btn				= 1700
	cm_04_12_btn					= 1800
	ebt_card_info_btn 	= 1900


	exp_page_display = show_exp_pg_amounts

	If first_time_in_exp_det = True Then
		If question_9_yn = "Yes" Then jobs_income_yn = "Yes"
		If question_9_yn = "No" Then jobs_income_yn = "No"
		If question_10_yn = "Yes" Then busi_income_yn = "Yes"
		If question_10_yn = "No" Then busi_income_yn = "No"
		exp_job_count = 0
		For each_caf_job = 0 to UBound(JOBS_ARRAY, 2)
			If JOBS_ARRAY(jobs_employer_name, each_caf_job) <> "" Then
				ReDim Preserve EXP_JOBS_ARRAY(jobs_notes_const, exp_job_count)
				EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count) = JOBS_ARRAY(jobs_employee_name, each_caf_job)
				If len(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count)) > 5 Then EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count) = right(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count), len(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count))-5)

				EXP_JOBS_ARRAY(jobs_employer_const, exp_job_count) = JOBS_ARRAY(jobs_employer_name, each_caf_job)
				EXP_JOBS_ARRAY(jobs_wage_const, exp_job_count) = JOBS_ARRAY(jobs_hourly_wage, each_caf_job)

				If IsNumeric(JOBS_ARRAY(jobs_gross_monthly_earnings, each_caf_job)) = True and IsNumeric(JOBS_ARRAY(jobs_hourly_wage, each_caf_job)) = True Then
                    If JOBS_ARRAY(jobs_hourly_wage, each_caf_job) > 0 Then      'making sure we are not dividing by zero. I will not be defaulting to a zero income job - no autofils
    					monthly_hours = JOBS_ARRAY(jobs_gross_monthly_earnings, each_caf_job)/JOBS_ARRAY(jobs_hourly_wage, each_caf_job)
    					weekly_hours = monthly_hours/4
    					EXP_JOBS_ARRAY(jobs_hours_const, exp_job_count) = weekly_hours
    					EXP_JOBS_ARRAY(jobs_frequency_const, exp_job_count) = "Weekly"
                    End If
				End If

				exp_job_count = exp_job_count + 1
			End If
		Next
		exp_unea_count = 0
		If IsNumeric(question_12_rsdi_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "RSDI"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_rsdi_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_ssi_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "SSI"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_ssi_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_va_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "VA Benefit"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_va_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_ui_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "Unemployment"
			EXP_UNEA_ARRAY(unea_weekly_earnings_const, exp_unea_count) = question_12_ui_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_wc_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "Workers Comp"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_wc_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_ret_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "Retirement Benefits"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_ret_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_trib_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "Tribal Payment"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_trib_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_cs_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = "Child Support"
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_cs_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If IsNumeric(question_12_other_amt) = True Then
			ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
			EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = ""
			EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = question_12_other_amt
			exp_unea_count = exp_unea_count + 1
		End If
		If exp_unea_count > 0 Then unea_income_yn = "Yes"

		Call app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, EXP_JOBS_ARRAY, EXP_BUSI_ARRAY, EXP_UNEA_ARRAY)

		If question_20_cash_yn = "Yes" Then cash_amount_yn = "Yes"
		If question_20_acct_yn = "Yes" Then bank_account_yn = "Yes"
		If question_20_cash_yn = "No" Then cash_amount_yn = "No"
		If question_20_acct_yn = "No" Then bank_account_yn = "No"
		Call app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, EXP_ACCT_ARRAY)


		If question_14_rent_yn = "Yes" Then
			rent_amount = exp_q_3_rent_this_month
		ElseIf question_14_mortgage_yn = "Yes" Then
			mortgage_amount = exp_q_3_rent_this_month
		ElseIf question_14_room_yn = "Yes" Then
			room_amount = exp_q_3_rent_this_month
		ElseIf question_14_insurance_yn = "Yes" Then
			insurance_amount = exp_q_3_rent_this_month
		ElseIf question_14_taxes_yn = "Yes" Then
			tax_amount = exp_q_3_rent_this_month
		End If


		Call app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)


		heat_expense = False
		ac_expense = False
		electric_expense = False
		phone_expense = False

		If question_15_heat_ac_yn = "Yes" Then
			heat_expense = True
			ac_expense = True
		End If
		If question_15_electricity_yn = "Yes" Then electric_expense = True
		If question_15_phone_yn = "Yes" Then phone_expense = True

		determined_utilities = 0
		If heat_expense = True OR ac_expense = True Then
			determined_utilities = determined_utilities + heat_AC_amt
		Else
			If electric_expense = True Then determined_utilities = determined_utilities + electric_amt
			If phone_expense = True Then determined_utilities = determined_utilities + phone_amt
		End If

		all_utilities = ""
		If heat_expense = True Then all_utilities = all_utilities & ", Heat"
		If ac_expense = True Then all_utilities = all_utilities & ", AC"
		If electric_expense = True Then all_utilities = all_utilities & ", Electric"
		If phone_expense = True Then all_utilities = all_utilities & ", Phone"
		If heat_expense = False AND ac_expense = False AND electric_expense = False AND phone_expense = False Then all_utilities = all_utilities & ", None"
		If left(all_utilities, 2) = ", " Then all_utilities = right(all_utilities, len(all_utilities) - 2)

		If question_15_heat_ac_yn = "" AND question_15_electricity_yn = "" AND question_15_phone_yn = "" Then Call app_month_utility_detail(determined_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, all_utilities)


		first_time_in_exp_det = False
	End If


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
				' End If
				' If exp_screening_note_found = False Then
				' 	Text 10, 20, 350, 10, "CASE:NOTE for Expedited Screening could not be found. No information to Display."
				' 	Text 10, 30, 350, 10, "Review Application for screening answers"
				' End If
				Text 10, 90, 370, 15, "Review and update the INCOME, ASSETS, and HOUSING EXPENSES as determined in the Interview."
				GroupBox 5, 105, 390, 125, "Information about Income, Resources, and Expenses"
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
				' If snap_elig_results_read = True Then Text 55, 200, 180, 10, "Autofilled information based on current STAT and ELIG panels"
				Text 15, 215, 250, 10, "Blank amounts will be defaulted to ZERO."
				' GroupBox 5, 220, 390, 100, "Supports"
				' Text 15, 235, 260, 10, "If you need support in handling for expedited, please access these resources:"
				' PushButton 25, 250, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
				' PushButton 25, 265, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
				' PushButton 25, 280, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
				' PushButton 25, 295, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
				' PushButton 180, 250, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
				' PushButton 180, 265, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
				' PushButton 180, 280, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn
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
			GroupBox 5, 295, 470, 60, "If you need support in handling for expedited, please access these resources:"
			PushButton 15, 305, 150, 13, "HSR Manual - Expedited SNAP", hsr_manual_expedited_snap_btn
			PushButton 15, 320, 150, 13, "HSR Manual - SNAP Applications", hsr_snap_applications_btn
			PushButton 15, 335, 150, 13, "SIR - SNAP Expedited Flowchart", sir_exp_flowchart_btn
			PushButton 165, 305, 150, 13, "Retrain Your Brain - Expedited - Identity", ryb_exp_identity_btn
			PushButton 165, 320, 150, 13, "Retrain Your Brain - Expedited - Timeliness", ryb_exp_timeliness_btn
			PushButton 315, 305, 150, 13, "CM 04.04 - SNAP / Expedited Food", cm_04_04_btn
			PushButton 315, 320, 150, 13, "CM 04.06 - 1st Mont Processing", cm_04_06_btn

			If exp_page_display <> show_exp_pg_amounts then PushButton 485, 10, 65, 13, "Amounts", amounts_btn
			If exp_page_display <> show_exp_pg_determination then PushButton 485, 25, 65, 13, "Determination", determination_btn
			If exp_page_display <> show_exp_pg_review then PushButton 485, 40, 65, 13, "Review", review_btn
			If exp_page_display <> show_exp_pg_review then PushButton 500, 365, 50, 15, "Next", next_btn
			If exp_page_display = show_exp_pg_review then PushButton 500, 365, 50, 15, "Return", finish_btn
			' CancelButton 500, 365, 50, 15
			' OkButton 500, 350, 50, 15
		EndDialog

		Dialog Dialog1

		' cancel_confirmation
		' MsgBox "1 - ButtonPressed is " & ButtonPressed

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

		If ButtonPressed = knowledge_now_support_btn Then Call send_support_email_to_KN
		If ButtonPressed = te_02_10_01_btn Then Call view_poli_temp("02", "10", "01", "")

		' MsgBox "2 - ButtonPressed is " & ButtonPressed

		' If page_display = show_exp_pg_amounts Then
		'
		' End If
		If exp_page_display = show_exp_pg_determination Then
			delay_due_to_interview = False
			do_we_have_applicant_id = "UNKNOWN"
			If applicant_id_on_file_yn = "Yes" OR applicant_id_through_SOLQ = "Yes" Then do_we_have_applicant_id = True
			If applicant_id_on_file_yn = "No" AND applicant_id_through_SOLQ = "No" Then do_we_have_applicant_id = False

			' If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbCr & "* The date of application needs to be entered as a valid date."
			' If IsDate(interview_date) = False Then err_msg = err_msg & vbCr & "* The interview date needs to be entered as a valid date. An Expedited Determination cannot be completed without the interview."
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

		' MsgBox "3 - ButtonPressed is " & ButtonPressed


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
		' MsgBox "4 - ButtonPressed is " & ButtonPressed

		If ButtonPressed >= 1000 Then
			If ButtonPressed = hsr_manual_expedited_snap_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Expedited_SNAP.aspx"
			If ButtonPressed = hsr_snap_applications_btn Then resource_URL = "https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/SNAP_Applications.aspx"
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

function guide_through_app_month_income()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Questions to Guide Determination of Income in Month of Application "
	  Text 10, 5, 435, 10, "These questions will help you to guide the resident through understanding what income we need to count for the month of application."
	  Text 10, 20, 150, 10, "FIRST - Explain to the resident these things:"
	  Text 25, 30, 410, 10, "- Income in the App Month is used to determine if we can get your some SNAP benefits right away - an EXPEDITED Issuance."
	  Text 25, 40, 410, 10, "- We just need a best estimate of this income - it doesn't have to be exact. There is no penalty for getting this detail incorrect."
	  Text 25, 50, 410, 10, "- I can help you walk through your income sources."
	  Text 25, 60, 350, 10, "-  We need you to answer these questions to complete the interview for your application for SNAP benefits."
	  GroupBox 5, 75, 440, 105, "JOBS Income: For every Job in the Household"
	  Text 15, 90, 200, 10, "How many paychecks have you received in MM/YY so far?"
	  Text 30, 105, 170, 10, "How much were all of the checks for, before taxes?"
	  Text 15, 120, 215, 10, "How many paychecks do you still expect to receive in MM/YY?"
	  Text 30, 135, 225, 10, "How many hours a week did you or will you work for these checks?"
	  Text 30, 150, 120, 10, "What is your rate of pay per hour?"
	  Text 30, 165, 255, 10, "Do you get tips/commission/bonuses? How much do you expect those to be?"
	  GroupBox 5, 185, 440, 90, "BUSI Income: For each self employment in the Household"
	  Text 15, 200, 235, 10, "How much do you typically receive in a month of this self employment?"
	  Text 15, 215, 275, 10, "Is your self employment based on a contract or contracts? And how are they paid?"
	  Text 15, 230, 305, 10, "If this is hard to determine, how much to you make in any other period (year, week, quarter)?"
	  Text 30, 245, 200, 10, "Is this consistent over the period or from period to period?"
	  Text 30, 260, 115, 10, "If it is not, what are the variations?"
	  GroupBox 5, 280, 440, 45, "UNEA Income: For each other source of income in the Household"
	  Text 15, 295, 200, 10, "How often and how much do you receive from each source?"
	  Text 15, 310, 230, 10, "If this is irregular, what have you gotten for the past couple months?"
	  Text 5, 330, 380, 10, "After calculating all of these income questions, repeat the amount and each source and confirm that it seems close."
	  ButtonGroup ButtonPressed
	    PushButton 395, 330, 50, 15, "Return", return_btn
	EndDialog

	dialog Dialog1

end function

function split_phone_number_into_parts(phone_variable, phone_left, phone_mid, phone_right)
'This function is to take the information provided as a phone number and split it up into the 3 parts
    phone_variable = trim(phone_variable)
    If phone_variable <> "" Then
        phone_variable = replace(phone_variable, "(", "")						'formatting the phone variable to get rid of symbols and spaces
        phone_variable = replace(phone_variable, ")", "")
        phone_variable = replace(phone_variable, "-", "")
        phone_variable = replace(phone_variable, " ", "")
        phone_variable = trim(phone_variable)
        phone_left = left(phone_variable, 3)									'reading the certain sections of the variable for each part.
        phone_mid = mid(phone_variable, 4, 3)
        phone_right = right(phone_variable, 4)
        phone_variable = "(" & phone_left & ")" & phone_mid & "-" & phone_right
    End If
end function

function validate_footer_month_entry(footer_month, footer_year, err_msg_var, bullet_char)
'This function will asses the variables provided as the footer month and year to be sure it is correct.
    If IsNumeric(footer_month) = FALSE Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be a number, review and reenter the footer month information."
    Else
        footer_month = footer_month * 1
        If footer_month > 12 OR footer_month < 1 Then err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer month should be between 1 and 12. Review and reenter the footer month information."
        footer_month = right("00" & footer_month, 2)
    End If

    If len(footer_year) < 2 Then
        err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be at least 2 characters long, review and reenter the footer year information."
    Else
        If IsNumeric(footer_year) = FALSE Then
            err_msg_var = err_msg_var & vbNewLine & bullet_char & " The footer year should be a number, review and reenter the footer year information."
        Else
            footer_year = right("00" & footer_year, 2)
        End If
    End If
end function

function save_your_work()
'This function records the variables into a txt file so that it can be retrieved by the script if run later.

	'Now determines name of file
	If MAXIS_case_number <> "" Then
		save_your_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
		If user_ID_for_validation = "ERHO003" Then save_your_work_path = user_c_drive_docs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
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

			objTextStream.WriteLine "01A - " & question_1_yn
			objTextStream.WriteLine "01N - " & question_1_notes
			objTextStream.WriteLine "01V - " & question_1_verif_yn
			objTextStream.WriteLine "01D - " & question_1_verif_details
			objTextStream.WriteLine "01I - " & question_1_interview_notes

			objTextStream.WriteLine "02A - " & question_2_yn
			objTextStream.WriteLine "02N - " & question_2_notes
			objTextStream.WriteLine "02V - " & question_2_verif_yn
			objTextStream.WriteLine "02D - " & question_2_verif_details
			objTextStream.WriteLine "02I - " & question_2_interview_notes

			objTextStream.WriteLine "03A - " & question_3_yn
			objTextStream.WriteLine "03N - " & question_3_notes
			objTextStream.WriteLine "03V - " & question_3_verif_yn
			objTextStream.WriteLine "03D - " & question_3_verif_details
			objTextStream.WriteLine "03I - " & question_3_interview_notes

			objTextStream.WriteLine "04A - " & question_4_yn
			objTextStream.WriteLine "04N - " & question_4_notes
			objTextStream.WriteLine "04V - " & question_4_verif_yn
			objTextStream.WriteLine "04D - " & question_4_verif_details
			objTextStream.WriteLine "04I - " & question_4_interview_notes

			objTextStream.WriteLine "05A - " & question_5_yn
			objTextStream.WriteLine "05N - " & question_5_notes
			objTextStream.WriteLine "05V - " & question_5_verif_yn
			objTextStream.WriteLine "05D - " & question_5_verif_details
			objTextStream.WriteLine "05I - " & question_5_interview_notes

			objTextStream.WriteLine "06A - " & question_6_yn
			objTextStream.WriteLine "06N - " & question_6_notes
			objTextStream.WriteLine "06V - " & question_6_verif_yn
			objTextStream.WriteLine "06D - " & question_6_verif_details
			objTextStream.WriteLine "06I - " & question_6_interview_notes

			objTextStream.WriteLine "07A - " & question_7_yn
			objTextStream.WriteLine "07N - " & question_7_notes
			objTextStream.WriteLine "07V - " & question_7_verif_yn
			objTextStream.WriteLine "07D - " & question_7_verif_details
			objTextStream.WriteLine "07I - " & question_7_interview_notes

			objTextStream.WriteLine "08A - " & question_8_yn
			objTextStream.WriteLine "08N - " & question_8_notes
			objTextStream.WriteLine "08V - " & question_8_verif_yn
			objTextStream.WriteLine "08D - " & question_8_verif_details
			objTextStream.WriteLine "08I - " & question_8_interview_notes

			objTextStream.WriteLine "09A - " & question_9_yn
			objTextStream.WriteLine "09N - " & question_9_notes
			objTextStream.WriteLine "09V - " & question_9_verif_yn
			objTextStream.WriteLine "09D - " & question_9_verif_details

			objTextStream.WriteLine "10A - " & question_10_yn
			objTextStream.WriteLine "10N - " & question_10_notes
			objTextStream.WriteLine "10V - " & question_10_verif_yn
			objTextStream.WriteLine "10D - " & question_10_verif_details
			objTextStream.WriteLine "10G - " & question_10_monthly_earnings
			objTextStream.WriteLine "10I - " & question_10_interview_notes

			objTextStream.WriteLine "11A - " & question_11_yn
			objTextStream.WriteLine "11N - " & question_11_notes
			objTextStream.WriteLine "11V - " & question_11_verif_yn
			objTextStream.WriteLine "11D - " & question_11_verif_details
			objTextStream.WriteLine "11I - " & question_11_interview_notes

			objTextStream.WriteLine "PWE - " & pwe_selection

			objTextStream.WriteLine "12A - RS - " & question_12_rsdi_yn
			objTextStream.WriteLine "12$ - RS - " & question_12_rsdi_amt
			objTextStream.WriteLine "12A - SS - " & question_12_ssi_yn
			objTextStream.WriteLine "12$ - SS - " & question_12_ssi_amt
			objTextStream.WriteLine "12A - VA - " & question_12_va_yn
			objTextStream.WriteLine "12$ - VA - " & question_12_va_amt
			objTextStream.WriteLine "12A - UI - " & question_12_ui_yn
			objTextStream.WriteLine "12$ - UI - " & question_12_ui_amt
			objTextStream.WriteLine "12A - WC - " & question_12_wc_yn
			objTextStream.WriteLine "12$ - WC - " & question_12_wc_amt
			objTextStream.WriteLine "12A - RT - " & question_12_ret_yn
			objTextStream.WriteLine "12$ - RT - " & question_12_ret_amt
			objTextStream.WriteLine "12A - TP - " & question_12_trib_yn
			objTextStream.WriteLine "12$ - TP - " & question_12_trib_amt
			objTextStream.WriteLine "12A - CS - " & question_12_cs_yn
			objTextStream.WriteLine "12$ - CS - " & question_12_cs_amt
			objTextStream.WriteLine "12A - OT - " & question_12_other_yn
			objTextStream.WriteLine "12$ - OT - " & question_12_other_amt
			objTextStream.WriteLine "12A - " & q_12_answered
			objTextStream.WriteLine "12N - " & question_12_notes
			objTextStream.WriteLine "12V - " & question_12_verif_yn
			objTextStream.WriteLine "12D - " & question_12_verif_details
			objTextStream.WriteLine "12I - " & question_12_interview_notes

			objTextStream.WriteLine "13A - " & question_13_yn
			objTextStream.WriteLine "13N - " & question_13_notes
			objTextStream.WriteLine "13V - " & question_13_verif_yn
			objTextStream.WriteLine "13D - " & question_13_verif_details
			objTextStream.WriteLine "13I - " & question_13_interview_notes

			objTextStream.WriteLine "14A - RT - " &  question_14_rent_yn
			objTextStream.WriteLine "14A - SB - " &  question_14_subsidy_yn
			objTextStream.WriteLine "14A - MT - " &  question_14_mortgage_yn
			objTextStream.WriteLine "14A - AS - " &  question_14_association_yn
			objTextStream.WriteLine "14A - IN - " &  question_14_insurance_yn
			objTextStream.WriteLine "14A - RM - " &  question_14_room_yn
			objTextStream.WriteLine "14A - TX - " &  question_14_taxes_yn
			objTextStream.WriteLine "14A - " & q_14_answered
			objTextStream.WriteLine "14N - " & question_14_notes
			objTextStream.WriteLine "14V - " & question_14_verif_yn
			objTextStream.WriteLine "14D - " & question_14_verif_details
			objTextStream.WriteLine "14I - " & question_14_interview_notes

			objTextStream.WriteLine "15A - HA - " & question_15_heat_ac_yn
			objTextStream.WriteLine "15A - EL - " & question_15_electricity_yn
			objTextStream.WriteLine "15A - CF - " & question_15_cooking_fuel_yn
			objTextStream.WriteLine "15A - WS - " & question_15_water_and_sewer_yn
			objTextStream.WriteLine "15A - GR - " & question_15_garbage_yn
			objTextStream.WriteLine "15A - PN - " & question_15_phone_yn
			objTextStream.WriteLine "15A - LP - " & question_15_liheap_yn
			objTextStream.WriteLine "15A - " & q_15_answered
			objTextStream.WriteLine "15N - " & question_15_notes
			objTextStream.WriteLine "15V - " & question_15_verif_yn
			objTextStream.WriteLine "15D - " & question_15_verif_details
			objTextStream.WriteLine "15I - " & question_15_interview_notes
			objTextStream.WriteLine "15PD - " & question_15_phone_details

			objTextStream.WriteLine "16A - " & question_16_yn
			objTextStream.WriteLine "16N - " & question_16_notes
			objTextStream.WriteLine "16V - " & question_16_verif_yn
			objTextStream.WriteLine "16D - " & question_16_verif_details
			objTextStream.WriteLine "16I - " & question_16_interview_notes

			objTextStream.WriteLine "17A - " & question_17_yn
			objTextStream.WriteLine "17N - " & question_17_notes
			objTextStream.WriteLine "17V - " & question_17_verif_yn
			objTextStream.WriteLine "17D - " & question_17_verif_details
			objTextStream.WriteLine "17I - " & question_17_interview_notes

			objTextStream.WriteLine "18A - " & question_18_yn
			objTextStream.WriteLine "18N - " & question_18_notes
			objTextStream.WriteLine "18V - " & question_18_verif_yn
			objTextStream.WriteLine "18D - " & question_18_verif_details
			objTextStream.WriteLine "18I - " & question_18_interview_notes

			objTextStream.WriteLine "19A - " & question_19_yn
			objTextStream.WriteLine "19N - " & question_19_notes
			objTextStream.WriteLine "19V - " & question_19_verif_yn
			objTextStream.WriteLine "19D - " & question_19_verif_details
			objTextStream.WriteLine "19I - " & question_19_interview_notes

			objTextStream.WriteLine "20A - CA - " & question_20_cash_yn
			objTextStream.WriteLine "20A - AC - " & question_20_acct_yn
			objTextStream.WriteLine "20A - SE - " & question_20_secu_yn
			objTextStream.WriteLine "20A - CR - " & question_20_cars_yn
			objTextStream.WriteLine "20A - " & q_20_answered
			objTextStream.WriteLine "20N - " & question_20_notes
			objTextStream.WriteLine "20V - " & question_20_verif_yn
			objTextStream.WriteLine "20D - " & question_20_verif_details
			objTextStream.WriteLine "20I - " & question_20_interview_notes

			objTextStream.WriteLine "21A - " & question_21_yn
			objTextStream.WriteLine "21N - " & question_21_notes
			objTextStream.WriteLine "21V - " & question_21_verif_yn
			objTextStream.WriteLine "21D - " & question_21_verif_details
			objTextStream.WriteLine "21I - " & question_21_interview_notes

			objTextStream.WriteLine "22A - " & question_22_yn
			objTextStream.WriteLine "22N - " & question_22_notes
			objTextStream.WriteLine "22V - " & question_22_verif_yn
			objTextStream.WriteLine "22D - " & question_22_verif_details
			objTextStream.WriteLine "22I - " & question_22_interview_notes

			objTextStream.WriteLine "23A - " & question_23_yn
			objTextStream.WriteLine "23N - " & question_23_notes
			objTextStream.WriteLine "23V - " & question_23_verif_yn
			objTextStream.WriteLine "23D - " & question_23_verif_details
			objTextStream.WriteLine "23I - " & question_23_interview_notes

			objTextStream.WriteLine "24A - RP - " & question_24_rep_payee_yn
			objTextStream.WriteLine "24A - GF - " & question_24_guardian_fees_yn
			objTextStream.WriteLine "24A - SD - " & question_24_special_diet_yn
			objTextStream.WriteLine "24A - HH - " & question_24_high_housing_yn
			objTextStream.WriteLine "24A - " & q_24_answered
			objTextStream.WriteLine "24N - " & question_24_notes
			objTextStream.WriteLine "24V - " & question_24_verif_yn
			objTextStream.WriteLine "24D - " & question_24_verif_details
			objTextStream.WriteLine "24I - " & question_24_interview_notes

			objTextStream.WriteLine "QQ1A - " & qual_question_one
			objTextStream.WriteLine "QQ1M - " & qual_memb_one
			objTextStream.WriteLine "QQ2A - " & qual_question_two
			objTextStream.WriteLine "QQ2M - " & qual_memb_two
			objTextStream.WriteLine "QQ3A - " & qual_question_three
			objTextStream.WriteLine "QQ3M - " & qual_memb_there
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

			objTextStream.WriteLine "FORM - 01 - " & confirm_resp_read
			objTextStream.WriteLine "FORM - 02 - " & confirm_rights_read
			objTextStream.WriteLine "FORM - 03 - " & confirm_ebt_read
			objTextStream.WriteLine "FORM -a03 - " & case_card_info
			objTextStream.WriteLine "FORM -b03 - " & clt_knows_how_to_use_ebt_card
			objTextStream.WriteLine "FORM - 04 - " & confirm_ebt_how_to_read
			objTextStream.WriteLine "FORM - 05 - " & confirm_npp_info_read
			objTextStream.WriteLine "FORM - 06 - " & confirm_npp_rights_read
			objTextStream.WriteLine "FORM - 07 - " & confirm_appeal_rights_read
			objTextStream.WriteLine "FORM - 08 - " & confirm_civil_rights_read
			objTextStream.WriteLine "FORM - 09 - " & confirm_cover_letter_read
			objTextStream.WriteLine "FORM - 10 - " & confirm_program_information_read
			objTextStream.WriteLine "FORM - 11 - " & confirm_DV_read
			objTextStream.WriteLine "FORM - 12 - " & confirm_disa_read
			objTextStream.WriteLine "FORM - 13 - " & confirm_mfip_forms_read
			objTextStream.WriteLine "FORM - 14 - " & confirm_mfip_cs_read
			objTextStream.WriteLine "FORM - 15 - " & confirm_minor_mfip_read
			objTextStream.WriteLine "FORM - 16 - " & confirm_snap_forms_read
			objTextStream.WriteLine "FORM -a16 - " & snap_reporting_type
			objTextStream.WriteLine "FORM -b16 - " & next_revw_month
			objTextStream.WriteLine "FORM - 17 - " & confirm_recap_read
			objTextStream.WriteLine "FORM - 18 - " & confirm_ievs_info_read

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

			for this_jobs = 0 to UBOUND(JOBS_ARRAY, 2)
				objTextStream.WriteLine "ARR - JOBS_ARRAY - " & JOBS_ARRAY(jobs_employee_name, this_jobs)&"~"&JOBS_ARRAY(jobs_hourly_wage, this_jobs)&"~"&JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)&"~"&_
				JOBS_ARRAY(jobs_employer_name, this_jobs)&"~"&JOBS_ARRAY(jobs_edit_btn, this_jobs)&"~"&JOBS_ARRAY(jobs_intv_notes, this_jobs)&"~"&JOBS_ARRAY(verif_yn, this_jobs)&"~"&JOBS_ARRAY(verif_details, this_jobs)&"~"&JOBS_ARRAY(jobs_notes, this_jobs)
			Next

			'Close the object so it can be opened again shortly
			objTextStream.Close

			script_run_lowdown = ""
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

			script_run_lowdown = script_run_lowdown & vbCr & "01A - " & question_1_yn
			script_run_lowdown = script_run_lowdown & vbCr & "01N - " & question_1_notes
			script_run_lowdown = script_run_lowdown & vbCr & "01V - " & question_1_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "01D - " & question_1_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "01I - " & question_1_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "02A - " & question_2_yn
			script_run_lowdown = script_run_lowdown & vbCr & "02N - " & question_2_notes
			script_run_lowdown = script_run_lowdown & vbCr & "02V - " & question_2_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "02D - " & question_2_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "02I - " & question_2_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "03A - " & question_3_yn
			script_run_lowdown = script_run_lowdown & vbCr & "03N - " & question_3_notes
			script_run_lowdown = script_run_lowdown & vbCr & "03V - " & question_3_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "03D - " & question_3_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "03I - " & question_3_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "04A - " & question_4_yn
			script_run_lowdown = script_run_lowdown & vbCr & "04N - " & question_4_notes
			script_run_lowdown = script_run_lowdown & vbCr & "04V - " & question_4_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "04D - " & question_4_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "04I - " & question_4_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "05A - " & question_5_yn
			script_run_lowdown = script_run_lowdown & vbCr & "05N - " & question_5_notes
			script_run_lowdown = script_run_lowdown & vbCr & "05V - " & question_5_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "05D - " & question_5_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "05I - " & question_5_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "06A - " & question_6_yn
			script_run_lowdown = script_run_lowdown & vbCr & "06N - " & question_6_notes
			script_run_lowdown = script_run_lowdown & vbCr & "06V - " & question_6_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "06D - " & question_6_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "06I - " & question_6_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "07A - " & question_7_yn
			script_run_lowdown = script_run_lowdown & vbCr & "07N - " & question_7_notes
			script_run_lowdown = script_run_lowdown & vbCr & "07V - " & question_7_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "07D - " & question_7_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "07I - " & question_7_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "08A - " & question_8_yn
			script_run_lowdown = script_run_lowdown & vbCr & "08N - " & question_8_notes
			script_run_lowdown = script_run_lowdown & vbCr & "08V - " & question_8_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "08D - " & question_8_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "08I - " & question_8_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "09A - " & question_9_yn
			script_run_lowdown = script_run_lowdown & vbCr & "09N - " & question_9_notes
			script_run_lowdown = script_run_lowdown & vbCr & "09V - " & question_9_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "09D - " & question_9_verif_details & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "10A - " & question_10_yn
			script_run_lowdown = script_run_lowdown & vbCr & "10N - " & question_10_notes
			script_run_lowdown = script_run_lowdown & vbCr & "10V - " & question_10_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "10D - " & question_10_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "10G - " & question_10_monthly_earnings
			script_run_lowdown = script_run_lowdown & vbCr & "10I - " & question_10_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "11A - " & question_11_yn
			script_run_lowdown = script_run_lowdown & vbCr & "11N - " & question_11_notes
			script_run_lowdown = script_run_lowdown & vbCr & "11V - " & question_11_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "11D - " & question_11_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "11I - " & question_11_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "PWE - " & pwe_selection & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "12A - RS - " & question_12_rsdi_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - RS - " & question_12_rsdi_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - SS - " & question_12_ssi_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - SS - " & question_12_ssi_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - VA - " & question_12_va_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - VA - " & question_12_va_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - UI - " & question_12_ui_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - UI - " & question_12_ui_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - WC - " & question_12_wc_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - WC - " & question_12_wc_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - RT - " & question_12_ret_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - RT - " & question_12_ret_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - TP - " & question_12_trib_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - TP - " & question_12_trib_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - CS - " & question_12_cs_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - CS - " & question_12_cs_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - OT - " & question_12_other_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12$ - OT - " & question_12_other_amt
			script_run_lowdown = script_run_lowdown & vbCr & "12A - " & q_12_answered
			script_run_lowdown = script_run_lowdown & vbCr & "12N - " & question_12_notes
			script_run_lowdown = script_run_lowdown & vbCr & "12V - " & question_12_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "12D - " & question_12_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "12I - " & question_12_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "13A - " & question_13_yn
			script_run_lowdown = script_run_lowdown & vbCr & "13N - " & question_13_notes
			script_run_lowdown = script_run_lowdown & vbCr & "13V - " & question_13_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "13D - " & question_13_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "13I - " & question_13_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "14A - RT - " &  question_14_rent_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - SB - " &  question_14_subsidy_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - MT - " &  question_14_mortgage_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - AS - " &  question_14_association_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - IN - " &  question_14_insurance_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - RM - " &  question_14_room_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - TX - " &  question_14_taxes_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14A - " & q_14_answered
			script_run_lowdown = script_run_lowdown & vbCr & "14N - " & question_14_notes
			script_run_lowdown = script_run_lowdown & vbCr & "14V - " & question_14_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "14D - " & question_14_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "14I - " & question_14_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "15A - HA - " & question_15_heat_ac_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - EL - " & question_15_electricity_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - CF - " & question_15_cooking_fuel_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - WS - " & question_15_water_and_sewer_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - GR - " & question_15_garbage_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - PN - " & question_15_phone_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - LP - " & question_15_liheap_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15A - " & q_15_answered
			script_run_lowdown = script_run_lowdown & vbCr & "15N - " & question_15_notes
			script_run_lowdown = script_run_lowdown & vbCr & "15V - " & question_15_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "15D - " & question_15_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "15I - " & question_15_interview_notes
			script_run_lowdown = script_run_lowdown & vbCr & "15PD - " & question_15_phone_details & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "16A - " & question_16_yn
			script_run_lowdown = script_run_lowdown & vbCr & "16N - " & question_16_notes
			script_run_lowdown = script_run_lowdown & vbCr & "16V - " & question_16_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "16D - " & question_16_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "16I - " & question_16_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "17A - " & question_17_yn
			script_run_lowdown = script_run_lowdown & vbCr & "17N - " & question_17_notes
			script_run_lowdown = script_run_lowdown & vbCr & "17V - " & question_17_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "17D - " & question_17_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "17I - " & question_17_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "18A - " & question_18_yn
			script_run_lowdown = script_run_lowdown & vbCr & "18N - " & question_18_notes
			script_run_lowdown = script_run_lowdown & vbCr & "18V - " & question_18_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "18D - " & question_18_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "18I - " & question_18_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "19A - " & question_19_yn
			script_run_lowdown = script_run_lowdown & vbCr & "19N - " & question_19_notes
			script_run_lowdown = script_run_lowdown & vbCr & "19V - " & question_19_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "19D - " & question_19_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "19I - " & question_19_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "20A - CA - " & question_20_cash_yn
			script_run_lowdown = script_run_lowdown & vbCr & "20A - AC - " & question_20_acct_yn
			script_run_lowdown = script_run_lowdown & vbCr & "20A - SE - " & question_20_secu_yn
			script_run_lowdown = script_run_lowdown & vbCr & "20A - CR - " & question_20_cars_yn
			script_run_lowdown = script_run_lowdown & vbCr & "20A - " & q_20_answered
			script_run_lowdown = script_run_lowdown & vbCr & "20N - " & question_20_notes
			script_run_lowdown = script_run_lowdown & vbCr & "20V - " & question_20_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "20D - " & question_20_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "20I - " & question_20_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "21A - " & question_21_yn
			script_run_lowdown = script_run_lowdown & vbCr & "21N - " & question_21_notes
			script_run_lowdown = script_run_lowdown & vbCr & "21V - " & question_21_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "21D - " & question_21_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "21I - " & question_21_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "22A - " & question_22_yn
			script_run_lowdown = script_run_lowdown & vbCr & "22N - " & question_22_notes
			script_run_lowdown = script_run_lowdown & vbCr & "22V - " & question_22_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "22D - " & question_22_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "22I - " & question_22_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "23A - " & question_23_yn
			script_run_lowdown = script_run_lowdown & vbCr & "23N - " & question_23_notes
			script_run_lowdown = script_run_lowdown & vbCr & "23V - " & question_23_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "23D - " & question_23_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "23I - " & question_23_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "24A - RP - " & question_24_rep_payee_yn
			script_run_lowdown = script_run_lowdown & vbCr & "24A - GF - " & question_24_guardian_fees_yn
			script_run_lowdown = script_run_lowdown & vbCr & "24A - SD - " & question_24_special_diet_yn
			script_run_lowdown = script_run_lowdown & vbCr & "24A - HH - " & question_24_high_housing_yn
			script_run_lowdown = script_run_lowdown & vbCr & "24A - " & q_24_answered
			script_run_lowdown = script_run_lowdown & vbCr & "24N - " & question_24_notes
			script_run_lowdown = script_run_lowdown & vbCr & "24V - " & question_24_verif_yn
			script_run_lowdown = script_run_lowdown & vbCr & "24D - " & question_24_verif_details
			script_run_lowdown = script_run_lowdown & vbCr & "24I - " & question_24_interview_notes & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "QQ1A - " & qual_question_one
			script_run_lowdown = script_run_lowdown & vbCr & "QQ1M - " & qual_memb_one
			script_run_lowdown = script_run_lowdown & vbCr & "QQ2A - " & qual_question_two
			script_run_lowdown = script_run_lowdown & vbCr & "QQ2M - " & qual_memb_two
			script_run_lowdown = script_run_lowdown & vbCr & "QQ3A - " & qual_question_three
			script_run_lowdown = script_run_lowdown & vbCr & "QQ3M - " & qual_memb_there
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

			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 01 - " & confirm_resp_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 02 - " & confirm_rights_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 03 - " & confirm_ebt_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -a03 - " & case_card_info
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -b03 - " & clt_knows_how_to_use_ebt_card
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 04 - " & confirm_ebt_how_to_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 05 - " & confirm_npp_info_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 06 - " & confirm_npp_rights_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 07 - " & confirm_appeal_rights_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 08 - " & confirm_civil_rights_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 09 - " & confirm_cover_letter_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 10 - " & confirm_program_information_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 11 - " & confirm_DV_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 12 - " & confirm_disa_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 13 - " & confirm_mfip_forms_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 14 - " & confirm_mfip_cs_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 15 - " & confirm_minor_mfip_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 16 - " & confirm_snap_forms_read
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -a16 - " & snap_reporting_type
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -b16 - " & next_revw_month
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 17 - " & confirm_recap_read & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 18 - " & confirm_ievs_info_read & vbCr & vbCr


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

			for this_jobs = 0 to UBOUND(JOBS_ARRAY, 2)
				script_run_lowdown = script_run_lowdown & vbCr & "ARR - JOBS_ARRAY - " & JOBS_ARRAY(jobs_employee_name, this_jobs)&"~"&JOBS_ARRAY(jobs_hourly_wage, this_jobs)&"~"&JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)&"~"&_
				JOBS_ARRAY(jobs_employer_name, this_jobs)&"~"&JOBS_ARRAY(jobs_edit_btn, this_jobs)&"~"&JOBS_ARRAY(jobs_intv_notes, this_jobs)&"~"&JOBS_ARRAY(verif_yn, this_jobs)&"~"&JOBS_ARRAY(verif_details, this_jobs)&"~"&JOBS_ARRAY(jobs_notes, this_jobs) & vbCr & vbCr
			Next

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

function restore_your_work(vars_filled)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run

	'Now determines name of file
	save_your_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
	If user_ID_for_validation = "ERHO003" Then save_your_work_path = user_c_drive_docs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"

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
					' If left(text_line, 3) = "" Then  = Mid(text_line, 7)
					If left(text_line, 3) = "01A" Then question_1_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01N" Then question_1_notes = Mid(text_line, 7)
					If left(text_line, 3) = "01V" Then question_1_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "01D" Then question_1_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "01I" Then question_1_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "02A" Then question_2_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02N" Then question_2_notes = Mid(text_line, 7)
					If left(text_line, 3) = "02V" Then question_2_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02D" Then question_2_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "02I" Then question_2_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "03A" Then question_3_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03N" Then question_3_notes = Mid(text_line, 7)
					If left(text_line, 3) = "03V" Then question_3_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03D" Then question_3_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "03I" Then question_3_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "04A" Then question_4_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04N" Then question_4_notes = Mid(text_line, 7)
					If left(text_line, 3) = "04V" Then question_4_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04D" Then question_4_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "04I" Then question_4_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "05A" Then question_5_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05N" Then question_5_notes = Mid(text_line, 7)
					If left(text_line, 3) = "05V" Then question_5_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05D" Then question_5_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "05I" Then question_5_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "06A" Then question_6_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06N" Then question_6_notes = Mid(text_line, 7)
					If left(text_line, 3) = "06V" Then question_6_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06D" Then question_6_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "06I" Then question_6_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "07A" Then question_7_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07N" Then question_7_notes = Mid(text_line, 7)
					If left(text_line, 3) = "07V" Then question_7_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07D" Then question_7_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "07I" Then question_7_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "08A" Then question_8_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08N" Then question_8_notes = Mid(text_line, 7)
					If left(text_line, 3) = "08V" Then question_8_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08D" Then question_8_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "08I" Then question_8_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "09A" Then question_9_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09N" Then question_9_notes = Mid(text_line, 7)
					If left(text_line, 3) = "09V" Then question_9_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09D" Then question_9_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "10A" Then question_10_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10N" Then question_10_notes = Mid(text_line, 7)
					If left(text_line, 3) = "10V" Then question_10_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10D" Then question_10_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "10G" Then question_10_monthly_earnings = Mid(text_line, 7)
					If left(text_line, 3) = "10I" Then question_10_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "11A" Then question_11_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11N" Then question_11_notes = Mid(text_line, 7)
					If left(text_line, 3) = "11V" Then question_11_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11D" Then question_11_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "11I" Then question_11_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "PWE" Then pwe_selection = Mid(text_line, 7)

					If left(text_line, 8) = "12A - RS" Then question_12_rsdi_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - RS" Then question_12_rsdi_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - SS" Then question_12_ssi_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - SS" Then question_12_ssi_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - VA" Then question_12_va_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - VA" Then question_12_va_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - UI" Then question_12_ui_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - UI" Then question_12_ui_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - WC" Then question_12_wc_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - WC" Then question_12_wc_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - RT" Then question_12_ret_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - RT" Then question_12_ret_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - TP" Then question_12_trib_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - TP" Then question_12_trib_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - CS" Then question_12_cs_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - CS" Then question_12_cs_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - OT" Then question_12_other_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - OT" Then question_12_other_amt = Mid(text_line, 12)
					If left(text_line, 3) = "12A" Then q_12_answered = Mid(text_line, 7)
					If left(text_line, 3) = "12N" Then question_12_notes = Mid(text_line, 7)
					If left(text_line, 3) = "12V" Then question_12_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "12D" Then question_12_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "12I" Then question_12_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "13A" Then question_13_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13N" Then question_13_notes = Mid(text_line, 7)
					If left(text_line, 3) = "13V" Then question_13_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13D" Then question_13_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "13I" Then question_13_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "14A - RT" Then  question_14_rent_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - SB" Then  question_14_subsidy_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - MT" Then  question_14_mortgage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - AS" Then  question_14_association_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - IN" Then  question_14_insurance_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - RM" Then  question_14_room_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - TX" Then  question_14_taxes_yn = Mid(text_line, 12)
					If left(text_line, 3) = "14A" Then q_14_answered = Mid(text_line, 7)
					If left(text_line, 3) = "14N" Then question_14_notes = Mid(text_line, 7)
					If left(text_line, 3) = "14V" Then question_14_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "14D" Then question_14_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "14I" Then question_14_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "15A - HA" Then question_15_heat_ac_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - EL" Then question_15_electricity_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - CF" Then question_15_cooking_fuel_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - WS" Then question_15_water_and_sewer_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - GR" Then question_15_garbage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - PN" Then question_15_phone_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - LP" Then question_15_liheap_yn = Mid(text_line, 12)
					If left(text_line, 3) = "15A" Then q_15_answered = Mid(text_line, 7)
					If left(text_line, 3) = "15N" Then question_15_notes = Mid(text_line, 7)
					If left(text_line, 3) = "15V" Then question_15_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "15D" Then question_15_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "15I" Then question_15_interview_notes = Mid(text_line, 7)
					If left(text_line, 4) = "15PD" Then question_15_phone_details = Mid(text_line, 8)

					If left(text_line, 3) = "16A" Then question_16_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16N" Then question_16_notes = Mid(text_line, 7)
					If left(text_line, 3) = "16V" Then question_16_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16D" Then question_16_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "16I" Then question_16_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "17A" Then question_17_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17N" Then question_17_notes = Mid(text_line, 7)
					If left(text_line, 3) = "17V" Then question_17_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17D" Then question_17_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "17I" Then question_17_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "18A" Then question_18_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18N" Then question_18_notes = Mid(text_line, 7)
					If left(text_line, 3) = "18V" Then question_18_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18D" Then question_18_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "18I" Then question_18_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "19A" Then question_19_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19N" Then question_19_notes = Mid(text_line, 7)
					If left(text_line, 3) = "19V" Then question_19_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19D" Then question_19_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "19I" Then question_19_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "20A - CA" Then question_20_cash_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - AC" Then question_20_acct_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - SE" Then question_20_secu_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - CR" Then question_20_cars_yn = Mid(text_line, 12)
					If left(text_line, 3) = "20A" Then q_20_answered = Mid(text_line, 7)
					If left(text_line, 3) = "20N" Then question_20_notes = Mid(text_line, 7)
					If left(text_line, 3) = "20V" Then question_20_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "20D" Then question_20_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "20I" Then question_20_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "21A" Then question_21_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21N" Then question_21_notes = Mid(text_line, 7)
					If left(text_line, 3) = "21V" Then question_21_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21D" Then question_21_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "21I" Then question_21_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "22A" Then question_22_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22N" Then question_22_notes = Mid(text_line, 7)
					If left(text_line, 3) = "22V" Then question_22_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22D" Then question_22_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "22I" Then question_22_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "23A" Then question_23_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23N" Then question_23_notes = Mid(text_line, 7)
					If left(text_line, 3) = "23V" Then question_23_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23D" Then question_23_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "23I" Then question_23_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "24A - RP" Then question_24_rep_payee_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - GF" Then question_24_guardian_fees_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - SD" Then question_24_special_diet_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - HH" Then question_24_high_housing_yn = Mid(text_line, 12)
					If left(text_line, 3) = "24A" Then q_24_answered = Mid(text_line, 7)
					If left(text_line, 3) = "24N" Then question_24_notes = Mid(text_line, 7)
					If left(text_line, 3) = "24V" Then question_24_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "24D" Then question_24_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "24I" Then question_24_interview_notes = Mid(text_line, 7)

					If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ1M" Then qual_memb_one = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2A" Then qual_question_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ2M" Then qual_memb_two = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3A" Then qual_question_three = Mid(text_line, 8)
					If left(text_line, 4) = "QQ3M" Then qual_memb_there = Mid(text_line, 8)
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

					If left(text_line, 9) = "FORM - 01" Then confirm_resp_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 02" Then confirm_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 03" Then confirm_ebt_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -a03" Then case_card_info = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -b03" Then clt_knows_how_to_use_ebt_card = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 04" Then confirm_ebt_how_to_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 05" Then confirm_npp_info_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 06" Then confirm_npp_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 07" Then confirm_appeal_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 08" Then confirm_civil_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 09" Then confirm_cover_letter_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 10" Then confirm_program_information_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 11" Then confirm_DV_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 12" Then confirm_disa_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 13" Then confirm_mfip_forms_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 14" Then confirm_mfip_cs_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 15" Then confirm_minor_mfip_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 16" Then confirm_snap_forms_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -a16" Then snap_reporting_type = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -b16" Then next_revw_month = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 17" Then confirm_recap_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 18" Then confirm_ievs_info_read = Mid(text_line, 13)
					' If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)

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

						If MID(text_line, 7, 10) = "JOBS_ARRAY" Then
							array_info = Mid(text_line, 20)
							array_info = split(array_info, "~")
							ReDim Preserve JOBS_ARRAY(jobs_notes, known_jobs)
							JOBS_ARRAY(jobs_employee_name, known_jobs) 			= array_info(0)
							JOBS_ARRAY(jobs_hourly_wage, known_jobs) 			= array_info(1)
							JOBS_ARRAY(jobs_gross_monthly_earnings, known_jobs)	= array_info(2)
							JOBS_ARRAY(jobs_employer_name, known_jobs) 			= array_info(3)
							JOBS_ARRAY(jobs_edit_btn, known_jobs)				= array_info(4)
							JOBS_ARRAY(jobs_intv_notes, known_jobs)				= array_info(5)
							JOBS_ARRAY(verif_yn, known_jobs)					= array_info(6)
							JOBS_ARRAY(verif_details, known_jobs)				= array_info(7)
							JOBS_ARRAY(jobs_notes, known_jobs) 					= array_info(8)
							known_jobs = known_jobs + 1
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

function review_information()
	for the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(id_verif, the_memb) = "Requested" Then
			If Instr(HH_MEMB_ARRAY(client_verification_details, the_memb), "Identity verification for M" & HH_MEMB_ARRAY(ref_number, the_memb) & " - " & HH_MEMB_ARRAY(full_name_const, the_memb)) = 0 Then
				HH_MEMB_ARRAY(client_verification, the_memb) = "Requested"
				If HH_MEMB_ARRAY(client_verification_details, the_memb) <> "" Then
					HH_MEMB_ARRAY(client_verification_details, the_memb) = HH_MEMB_ARRAY(client_verification_details, the_memb) & ", Identity verification for M" & HH_MEMB_ARRAY(ref_number, the_memb) & " - " & HH_MEMB_ARRAY(full_name_const, the_memb)
				Else
					HH_MEMB_ARRAY(client_verification_details, the_memb) = "Identity verification for M" & HH_MEMB_ARRAY(ref_number, the_memb) & " - " & HH_MEMB_ARRAY(full_name_const, the_memb)
				End If
			End If
		End If
		If HH_MEMB_ARRAY(full_name_const, the_memb) = "" Then HH_MEMB_ARRAY(full_name_const, the_memb) = HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb)
	next
	q_12_totally_blank = True
	If question_12_rsdi_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_rsdi_amt) <> "" Then q_12_totally_blank = False

	If question_12_ssi_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_ssi_amt) <> "" Then q_12_totally_blank = False

	If question_12_va_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_va_amt) <> "" Then q_12_totally_blank = False

	If question_12_ui_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_ui_amt) <> "" Then q_12_totally_blank = False

	If question_12_wc_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_wc_amt) <> "" Then q_12_totally_blank = False

	If question_12_ret_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_ret_amt) <> "" Then q_12_totally_blank = False

	If question_12_trib_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_trib_amt) <> "" Then q_12_totally_blank = False

	If question_12_cs_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_cs_amt) <> "" Then q_12_totally_blank = False

	If question_12_other_yn <> "" Then q_12_totally_blank = False
	If trim(question_12_other_amt) <> "" Then q_12_totally_blank = False
	If trim(question_12_notes) <> "" Then q_12_totally_blank = False

	q_14_totally_blank = True
	If question_14_rent_yn <> "" Then q_14_totally_blank = False
	If question_14_subsidy_yn <> "" Then q_14_totally_blank = False
	If question_14_mortgage_yn <> "" Then q_14_totally_blank = False
	If question_14_taxes_yn <> "" Then q_14_totally_blank = False
	If question_14_association_yn <> "" Then q_14_totally_blank = False
	If question_14_insurance_yn <> "" Then q_14_totally_blank = False
	If question_14_room_yn <> "" Then q_14_totally_blank = False
	If trim(question_14_notes) <> "" Then q_14_totally_blank = False

	q_15_totally_blank = True
	If question_15_heat_ac_yn <> "" Then q_15_totally_blank = False
	If question_15_electricity_yn <> "" Then q_15_totally_blank = False
	If question_15_cooking_fuel_yn <> "" Then q_15_totally_blank = False
	If question_15_water_and_sewer_yn <> "" Then q_15_totally_blank = False
	If question_15_garbage_yn <> "" Then q_15_totally_blank = False
	If question_15_phone_yn <> "" Then q_15_totally_blank = False
	If question_15_liheap_yn <> "" Then q_15_totally_blank = False
	If trim(question_15_notes) <> "" Then q_15_totally_blank = False

	q_20_totally_blank = True
	If question_20_cash_yn <> "" Then q_20_totally_blank = False
	If question_20_acct_yn <> "" Then q_20_totally_blank = False
	If question_20_secu_yn <> "" Then q_20_totally_blank = False
	If question_20_cars_yn <> "" Then q_20_totally_blank = False
	If trim(question_20_notes) <> "" Then q_20_totally_blank = False

	q_24_totally_blank = True
	If question_24_rep_payee_yn <> "" Then q_24_totally_blank = False
	If question_24_guardian_fees_yn <> "" Then q_24_totally_blank = False
	If question_24_special_diet_yn <> "" Then q_24_totally_blank = False
	If question_24_high_housing_yn <> "" Then q_24_totally_blank = False
	If trim(question_24_notes) <> "" Then q_24_totally_blank = False

end function

function review_for_discrepancies()

	'PHONE NUMBER
	phone_one_number = trim(phone_one_number)
	phone_two_number = trim(phone_two_number)
	phone_three_number = trim(phone_three_number)
	disc_phone_confirmation = trim(disc_phone_confirmation)

	If phone_one_number = "" AND phone_two_number = "" AND phone_three_number = "" Then disc_no_phone_number = "EXISTS"
	If phone_one_number <> "" OR phone_two_number <> "" OR phone_three_number <> "" Then disc_no_phone_number = "N/A"

	If disc_no_phone_number <> "N/A" Then
		If disc_phone_confirmation <> "" and disc_phone_confirmation <> "Select or Type" Then disc_no_phone_number = "RESOLVED"
	Else
		disc_phone_confirmation = ""
	End If

	'HOMELESS NO MAILING ADDRESS
	' mail_addr_street_full = trim(mail_addr_street_full)
	' resi_street_to_look_at = trim(resi_addr_street_full)
	' resi_street_to_look_at = UBound(resi_street_to_look_at)
	' resi_street_appears_general_delivery = False
	' If Instr(resi_street_to_look_at, "GENERAL DELIVERY") Then resi_street_appears_general_delivery = True
	' If Instr(resi_street_to_look_at, "GENERALDELIVERY") Then resi_street_appears_general_delivery = True
	' If Instr(resi_street_to_look_at, "GEN DELIVERY") Then resi_street_appears_general_delivery = True
	' If Instr(resi_street_to_look_at, "GENERAL DEL") Then resi_street_appears_general_delivery = True
	' If Instr(resi_street_to_look_at, "GEN DEL") Then resi_street_appears_general_delivery = True

	If homeless_yn = "Yes" Then disc_homeless_no_mail_addr = "EXISTS"
	If homeless_yn <> "Yes" Then disc_homeless_no_mail_addr = "N/A"

	' If mail_addr_street_full = "" and resi_street_appears_general_delivery = True Then disc_homeless_no_mail_addr = "EXISTS"
	' End If
	If disc_homeless_no_mail_addr <> "N/A" Then
		If disc_homeless_confirmation <> "" and disc_homeless_confirmation <> "Select or Type" Then disc_homeless_no_mail_addr = "RESOLVED"
	Else
		disc_homeless_confirmation = ""
	End If

	'PHONE NUMBER BUT NO PHONE EXPENSE
	disc_yes_phone_no_expense_confirmation = trim(disc_yes_phone_no_expense_confirmation)
	disc_no_phone_yes_expense_confirmation = trim(disc_no_phone_yes_expense_confirmation)
	question_15_phone_details = trim(question_15_phone_details)
	disc_yes_phone_no_expense = "N/A"
	disc_no_phone_yes_expense = "N/A"

	If phone_one_number <> "" OR phone_two_number <> "" OR phone_three_number <> "" Then
		If question_15_phone_yn <> "Yes" Then disc_yes_phone_no_expense = "EXISTS"
		If caf_exp_pay_phone_checkbox = unchecked Then disc_yes_phone_no_expense = "EXISTS"
	End If
	If phone_one_number = "" AND phone_two_number = "" AND phone_three_number = "" Then
		If question_15_phone_yn = "Yes" Then disc_no_phone_yes_expense = "EXISTS"
		If caf_exp_pay_phone_checkbox = checked Then disc_no_phone_yes_expense = "EXISTS"
	End If

	If disc_yes_phone_no_expense <> "N/A" Then
		If question_15_phone_details <> "" AND question_15_phone_details <> "Select or Type" Then disc_yes_phone_no_expense_confirmation = question_15_phone_details
		If disc_yes_phone_no_expense_confirmation <> "" and disc_yes_phone_no_expense_confirmation <> "Select or Type" Then disc_yes_phone_no_expense = "RESOLVED"
	Else
		disc_yes_phone_no_expense_confirmation = ""
	End If
	If disc_no_phone_yes_expense <> "N/A" Then
		If disc_no_phone_yes_expense_confirmation <> "" and disc_no_phone_yes_expense_confirmation <> "Select or Type" Then disc_no_phone_yes_expense = "RESOLVED"
	Else
		disc_no_phone_yes_expense_confirmation = ""
	End If

	'OUT OF COUNTY
	If left(resi_addr_county, 2) <> "27" Then disc_out_of_county = "EXISTS"
	If left(resi_addr_county, 2) = "27" Then disc_out_of_county = "N/A"

	If disc_out_of_county <> "N/A" Then
		If disc_out_of_county_confirmation <> "" and disc_out_of_county_confirmation <> "Select or Type" Then disc_out_of_county = "RESOLVED"
	Else
		disc_out_of_county_confirmation = ""
	End If

	'RENT AMOUNTS
	exp_q_3_rent_this_month = trim(exp_q_3_rent_this_month)
	CAF1_rent_indicated = True
	If exp_q_3_rent_this_month = "" Then
		CAF1_rent_indicated = False
	ElseIf exp_q_3_rent_this_month = "0" Then
		CAF1_rent_indicated = False
	ElseIf exp_q_3_rent_this_month = 0 Then
		CAF1_rent_indicated = False
	End If

	intv_app_month_housing_expense = trim(intv_app_month_housing_expense)
	Verbal_rent_indicated = True
	If intv_app_month_housing_expense = "" Then
		Verbal_rent_indicated = False
	ElseIf intv_app_month_housing_expense = "0" Then
		Verbal_rent_indicated = False
	ElseIf intv_app_month_housing_expense = 0 Then
		Verbal_rent_indicated = False
	End If

	Q14_rent_indicated = False
	question_14_summary = ""
	If question_14_rent_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Rent"
	End If
	If question_14_subsidy_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Subsidy"
	End If
	If question_14_mortgage_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Mortgage"
	End If
	If question_14_association_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Association Fees"
	End If
	If question_14_insurance_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Home Insurance"
	End If
	If question_14_room_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Room or Board"
	End If
	If question_14_taxes_yn = "Yes" Then
		Q14_rent_indicated = True
		question_14_summary = question_14_summary & "/Real Estate Taxes"
	End If
	If left(question_14_summary, 1) = "/" Then question_14_summary = right(question_14_summary, len(question_14_summary) - 1)
	If question_14_summary = "" Then question_14_summary = "None Indicated"

	If CAF1_rent_indicated <> Q14_rent_indicated Then disc_rent_amounts = "EXISTS"
	If CAF1_rent_indicated = Q14_rent_indicated Then disc_rent_amounts = "N/A"

	If disc_rent_amounts <> "N/A" Then
		If disc_rent_amounts_confirmation <> "" and disc_rent_amounts_confirmation <> "Select or Type" Then disc_rent_amounts = "RESOLVED"
	Else
		disc_rent_amounts_confirmation = ""
	End If

	'UTILITY AMOUNTS
	disc_utility_amounts = "N/A"
	If caf_exp_pay_heat_checkbox = checked AND question_15_heat_ac_yn <> "Yes" Then disc_utility_amounts = "EXISTS"
	If caf_exp_pay_ac_checkbox = checked AND question_15_heat_ac_yn <> "Yes" Then disc_utility_amounts = "EXISTS"
	If caf_exp_pay_electricity_checkbox = checked AND question_15_electricity_yn <> "Yes" Then disc_utility_amounts = "EXISTS"
	If caf_exp_pay_phone_checkbox = checked AND question_15_phone_yn <> "Yes" Then disc_utility_amounts = "EXISTS"
	If caf_exp_pay_none_checkbox = checked Then
		If question_15_heat_ac_yn = "Yes" Then disc_utility_amounts = "EXISTS"
		If question_15_electricity_yn = "Yes" Then disc_utility_amounts = "EXISTS"
		If question_15_phone_yn = "Yes" Then disc_utility_amounts = "EXISTS"
	End If
	disc_utility_caf_1_summary = ""
	If caf_exp_pay_heat_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Heat"
	If caf_exp_pay_ac_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", AC"
	If caf_exp_pay_electricity_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Electricity"
	If caf_exp_pay_phone_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Phone"
	If caf_exp_pay_none_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", NONE"
	If left(disc_utility_caf_1_summary, 1) = "," Then disc_utility_caf_1_summary = right(disc_utility_caf_1_summary, len(disc_utility_caf_1_summary) - 2)

	disc_utility_q_15_summary = ""
	If question_15_heat_ac_yn = "Yes" Then disc_utility_q_15_summary = disc_utility_q_15_summary & ", Heat/AC"
	If question_15_electricity_yn = "Yes" Then disc_utility_q_15_summary = disc_utility_q_15_summary & ", Electricity"
	If question_15_phone_yn = "Yes" Then disc_utility_q_15_summary = disc_utility_q_15_summary & ", Phone"
	If left(disc_utility_q_15_summary, 1) = "," Then disc_utility_q_15_summary = right(disc_utility_q_15_summary, len(disc_utility_q_15_summary) - 2)
	If disc_utility_q_15_summary = "" Then disc_utility_q_15_summary = "None Indicated"

	If disc_utility_amounts <> "N/A" Then
		If disc_utility_amounts_confirmation <> "" and disc_utility_amounts_confirmation <> "Select or Type" Then disc_utility_amounts = "RESOLVED"
	Else
		disc_utility_amounts_confirmation = ""
	End If

	If disc_no_phone_number <> "N/A" Then discrepancies_exist = True
	If disc_homeless_no_mail_addr <> "N/A" Then discrepancies_exist = True
	If disc_out_of_county <> "N/A" Then discrepancies_exist = True
	If disc_rent_amounts <> "N/A" Then discrepancies_exist = True
	If disc_utility_amounts <> "N/A" Then discrepancies_exist = True
	If disc_yes_phone_no_expense <> "N/A" Then discrepancies_exist = True
	If disc_no_phone_yes_expense <> "N/A" Then discrepancies_exist = True

	If disc_no_phone_number = "N/A" and disc_homeless_no_mail_addr = "N/A" and disc_out_of_county = "N/A" and disc_rent_amounts = "N/A" and disc_utility_amounts = "N/A" and disc_yes_phone_no_expense = "N/A" and disc_no_phone_yes_expense = "N/A" Then discrepancies_exist = False
end function

function verif_details_dlg(question_number)
	Select Case question_number
		Case 1
			verif_selection = question_1_verif_yn
			verif_detials = question_1_verif_details
			question_words = "1. Does everyone in your household buy, fix or eat food with you?"
		Case 2
			verif_selection = question_2_verif_yn
			verif_detials = question_2_verif_details
			question_words = "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		Case 3
			verif_selection = question_3_verif_yn
			verif_detials = question_3_verif_details
			question_words = "3. Is anyone in the household attending school?"
		Case 4
			verif_selection = question_4_verif_yn
			verif_detials = question_4_verif_details
			question_words = "4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)"
		Case 5
			verif_selection = question_5_verif_yn
			verif_detials = question_5_verif_details
			question_words = "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
		Case 6
			verif_selection = question_6_verif_yn
			verif_detials = question_6_verif_details
			question_words = "6. Is anyone unable to work for reasons other than illness or disability?"
		Case 7
			verif_selection = question_7_verif_yn
			verif_detials = question_7_verif_details
			question_words = "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
		Case 8
			verif_selection = question_8_verif_yn
			verif_detials = question_8_verif_details
			question_words = "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
		Case 9
			verif_selection = question_9_verif_yn
			verif_detials = question_9_verif_details
			question_words = "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
		Case 10
			verif_selection = question_10_verif_yn
			verif_detials = question_10_verif_details
			question_words = "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
		Case 11
			verif_selection = question_11_verif_yn
			verif_detials = question_11_verif_details
			question_words = "11. Do you expect any changes in income, expenses or work hours?"
		Case 12
			verif_selection = question_12_verif_yn
			verif_detials = question_12_verif_details
			question_words = "12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
		Case 13
			verif_selection = question_13_verif_yn
			verif_detials = question_13_verif_details
			question_words = "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
		Case 14
			verif_selection = question_14_verif_yn
			verif_detials = question_14_verif_details
			question_words = "14. Does your household have the following housing expenses?"
		Case 15
			verif_selection = question_15_verif_yn
			verif_detials = question_15_verif_details
			question_words = "15. Does your household have the following utility expenses any time during the year?"
		Case 16
			verif_selection = question_16_verif_yn
			verif_detials = question_16_verif_details
			question_words = "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
		Case 17
			verif_selection = question_17_verif_yn
			verif_detials = question_17_verif_details
			question_words = "17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
		Case 18
			verif_selection = question_18_verif_yn
			verif_detials = question_18_verif_details
			question_words = "18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
		Case 19
			verif_selection = question_19_verif_yn
			verif_detials = question_19_verif_details
			question_words = "19. For SNAP only: Does anyone in the household have medical expenses? "
		Case 20
			verif_selection = question_20_verif_yn
			verif_detials = question_20_verif_details
			question_words = "20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
		Case 21
			verif_selection = question_21_verif_yn
			verif_detials = question_21_verif_details
			question_words = "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)"
		Case 22
			verif_selection = question_22_verif_yn
			verif_detials = question_22_verif_details
			question_words = "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
		Case 23
			verif_selection = question_23_verif_yn
			verif_detials = question_23_verif_details
			question_words = "23. For children under the age of 19, are both parents living in the home?"
		Case 24
			verif_selection = question_24_verif_yn
			verif_detials = question_24_verif_details
			question_words = "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
		Case 25
			verif_selection = JOBS_ARRAY(verif_yn, this_jobs)
			verif_detials = JOBS_ARRAY(verif_details, this_jobs)
			question_words = "9.  Does anyone in the household have a job or expect to get income from a job this month or next month? Enter verification for "	& JOBS_ARRAY(jobs_employer_name, this_jobs)
	End Select


	BeginDialog Dialog1, 0, 0, 396, 95, "Add Verification"
	  DropListBox 60, 35, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", verif_selection
	  EditBox 60, 55, 330, 15, verif_detials
	  ButtonGroup ButtonPressed
	    PushButton 340, 75, 50, 15, "Return", return_btn
		PushButton 145, 35, 50, 10, "CLEAR", clear_btn
	  Text 10, 10, 380, 20, question_words
	  Text 10, 40, 45, 10, "Verification: "
	  Text 20, 60, 30, 10, "Details:"
	EndDialog

	Do
		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = clear_btn Then
			verif_selection = "Not Needed"
			verif_detials = ""
		End If
	Loop until ButtonPressed = return_btn

	Select Case question_number
		Case 1
			question_1_verif_yn = verif_selection
			question_1_verif_details = verif_detials
		Case 2
			question_2_verif_yn = verif_selection
			question_2_verif_details = verif_detials
		Case 3
			question_3_verif_yn = verif_selection
			question_3_verif_details = verif_detials
		Case 4
			question_4_verif_yn = verif_selection
			question_4_verif_details = verif_detials
		Case 5
			question_5_verif_yn = verif_selection
			question_5_verif_details = verif_detials
		Case 6
			question_6_verif_yn = verif_selection
			question_6_verif_details = verif_detials
		Case 7
			question_7_verif_yn = verif_selection
			question_7_verif_details = verif_detials
		Case 8
			question_8_verif_yn = verif_selection
			question_8_verif_details = verif_detials
		Case 9
			question_9_verif_yn = verif_selection
			question_9_verif_details = verif_detials
		Case 10
			question_10_verif_yn = verif_selection
			question_10_verif_details = verif_detials
		Case 11
			question_11_verif_yn = verif_selection
			question_11_verif_details = verif_detials
		Case 12
			question_12_verif_yn = verif_selection
			question_12_verif_details = verif_detials
		Case 13
			question_13_verif_yn = verif_selection
			question_13_verif_details = verif_detials
		Case 14
			question_14_verif_yn = verif_selection
			question_14_verif_details = verif_detials
		Case 15
			question_15_verif_yn = verif_selection
			question_15_verif_details = verif_detials
		Case 16
			question_16_verif_yn = verif_selection
			question_16_verif_details = verif_detials
		Case 17
			question_17_verif_yn = verif_selection
			question_17_verif_details = verif_detials
		Case 18
			question_18_verif_yn = verif_selection
			question_18_verif_details = verif_detials
		Case 19
			question_19_verif_yn = verif_selection
			question_19_verif_details = verif_detials
		Case 20
			question_20_verif_yn = verif_selection
			question_20_verif_details = verif_detials
		Case 21
			question_21_verif_yn = verif_selection
			question_21_verif_details = verif_detials
		Case 22
			question_22_verif_yn = verif_selection
			question_22_verif_details = verif_detials
		Case 23
			question_23_verif_yn = verif_selection
			question_23_verif_details = verif_detials
		Case 24
			question_24_verif_yn = verif_selection
			question_24_verif_details = verif_detials
		Case 25
			JOBS_ARRAY(verif_yn, this_jobs) = verif_selection
			JOBS_ARRAY(verif_details, this_jobs) = verif_detials
	End Select

end function

function jobs_details_dlg(this_jobs)
	Do

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 321, 165, "Add Job"
		  DropListBox 10, 35, 135, 45, pick_a_client+chr(9)+"", JOBS_ARRAY(jobs_employee_name, this_jobs)
		  EditBox 150, 35, 60, 15, JOBS_ARRAY(jobs_hourly_wage, this_jobs)
		  EditBox 215, 35, 100, 15, JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)
		  EditBox 10, 65, 305, 15, JOBS_ARRAY(jobs_employer_name, this_jobs)
		  EditBox 10, 95, 305, 15, JOBS_ARRAY(jobs_notes, this_jobs)
		  EditBox 10, 125, 305, 15, JOBS_ARRAY(jobs_intv_notes, this_jobs)

		  ButtonGroup ButtonPressed
		    PushButton 265, 145, 50, 15, "Return", return_btn
			PushButton 120, 150, 75, 10, "ADD VERIFICATION", add_verif_jobs_btn
		    PushButton 265, 10, 50, 10, "CLEAR", clear_job_btn
		  Text 10, 10, 100, 10, "Enter Job Details/Information"
		  Text 10, 25, 70, 10, "EMPLOYEE NAME:"
		  Text 150, 25, 60, 10, "HOURLY WAGE:"
		  Text 215, 25, 105, 10, "GROSS MONTHLY EARNINGS:"
		  Text 10, 55, 110, 10, "EMPLOYER/BUSINESS NAME:"
		  Text 10, 85, 110, 10, "CAF WRITE-IN INFORMATION:"
		  Text 10, 115, 85, 10, "INTERVIEW NOTES:"
		  Text 10, 150, 110, 10, "JOB Verification - " & JOBS_ARRAY(verif_yn, this_jobs)
		EndDialog


		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = add_verif_jobs_btn Then Call verif_details_dlg(25)
		If ButtonPressed = clear_job_btn Then
			JOBS_ARRAY(jobs_employee_name, this_jobs) = ""
			JOBS_ARRAY(jobs_hourly_wage, this_jobs) = ""
			JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs) = ""
			JOBS_ARRAY(jobs_employer_name, this_jobs) = ""
			JOBS_ARRAY(jobs_notes, this_jobs) = ""
		End If
	Loop until ButtonPressed = return_btn
	If JOBS_ARRAY(jobs_employee_name, this_jobs) = "Select One..." Then JOBS_ARRAY(jobs_employee_name, this_jobs) = ""

end function

function format_phone_number(phone_variable, format_type)
'This function formats phone numbers to match the specificed format.
	' format_type_options:
	'  (xxx)xxx-xxxx
	'  xxx-xxx-xxxx
	'  xxx xxx xxxx
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) = 10 Then
		left_phone = left(phone_variable, 3)
		mid_phone = mid(phone_variable, 4, 3)
		right_phone = right(phone_variable, 4)
		format_type = lcase(format_type)
		If format_type = "(xxx)xxx-xxxx" Then
			phone_variable = "(" & left_phone & ")" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx-xxx-xxxx" Then
			phone_variable = left_phone & "-" & mid_phone & "-" & right_phone
		End If
		If format_type = "xxx xxx xxxx" Then
			phone_variable = left_phone & " " & mid_phone & " " & right_phone
		End If
	Else
		phone_variable = original_phone_var
	End If
end function

function validate_phone_number(err_msg_variable, list_delimiter, phone_variable, allow_to_be_blank)
'This isn't working yet
'This function will review to ensure a variale appears to be a phone number.
	original_phone_var = phone_variable
	phone_variable = trim(phone_variable)
	phone_variable = replace(phone_variable, "(", "")
	phone_variable = replace(phone_variable, ")", "")
	phone_variable = replace(phone_variable, "-", "")
	phone_variable = replace(phone_variable, " ", "")

	If len(phone_variable) <> 10 Then err_msg_variable = err_msg_variable & vbNewLine & list_delimiter & " Phone numbers should be entered as a 10 digit number. Please incldue the area code or check the number to ensure the correct information is entered."
	If len(phone_variable) = 0 then
		If allow_to_be_blank = TRUE then err_msg_variable = ""
	End If
	phone_variable = original_phone_var
end function

function verification_dialog()
    If ButtonPressed = verif_button Then
        If second_call <> TRUE Then
            ' income_source_list = "Select or Type Source"

            ' For each_job = 0 to UBound(ALL_JOBS_PANELS_ARRAY, 2)
            '     If ALL_JOBS_PANELS_ARRAY(employer_name, each_job) <> "" Then income_source_list = income_source_list+chr(9)+"JOB - " & ALL_JOBS_PANELS_ARRAY(employer_name, each_job)
            ' Next
            ' For each_busi = 0 to UBound(ALL_BUSI_PANELS_ARRAY, 2)
            '     If ALL_BUSI_PANELS_ARRAY(memb_numb, each_busi) <> "" Then
            '         If ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi) <> "" Then
            '             income_source_list = income_source_list+chr(9)+"Self Emp - " & ALL_BUSI_PANELS_ARRAY(busi_desc, each_busi)
            '         Else
            '             income_source_list = income_source_list+chr(9)+"Self Employment"
            '         End If
            '     End If
            ' Next
            ' employment_source_list = income_source_list
            income_source_list = "Select or Type Source"+chr(9)+"Job"+chr(9)+"Self Employment"+chr(9)+"Child Support"+chr(9)+"Social Security Income"+chr(9)+"Unemployment Income"+chr(9)+"VA Income"+chr(9)+"Pension"
            income_verif_time = "[Enter Time Frame]"
            bank_verif_time = "[Enter Time Frame]"
            second_call = TRUE
        End If

        Do
            verif_err_msg = ""
			' BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"

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
				For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
					If HH_MEMB_ARRAY(client_verification, the_members) = "Requested" Then
						Text 10, y_pos, 500, 10, "MEMB " & HH_MEMB_ARRAY(ref_number, the_members) & "-" & HH_MEMB_ARRAY(full_name_const, the_members) & " Information. Details: " & HH_MEMB_ARRAY(client_verification_details, the_members)
						y_pos = y_pos + 15
						grp_len = grp_len + 15
					End If
				Next
			  	If question_1_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q1 Verif Requested. Details: " & question_1_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_2_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q2 Verif Requested. Details: " & question_2_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_3_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q3 Verif Requested. Details: " & question_3_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_4_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q4 Verif Requested. Details: " & question_4_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_5_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q5 Verif Requested. Details: " & question_5_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_6_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q6 Verif Requested. Details: " & question_6_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_7_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q7 Verif Requested. Details: " & question_7_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_8_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q8 Verif Requested. Details: " & question_8_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				For each_job = 0 to UBound(JOBS_ARRAY, 2)
					If JOBS_ARRAY(verif_yn, each_job) = "Requested" Then
						Text 10, y_pos, 500, 10, "Q9 Verif Requested for " & JOBS_ARRAY(employer_name, each_job) & ". Details: " & JOBS_ARRAY(verif_details, each_job)
						y_pos = y_pos + 15
						grp_len = grp_len + 15
					End If
				Next
				If question_10_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q10 Verif Requested. Details: " & question_10_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_11_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q11 Verif Requested. Details: " & question_11_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_12_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q12 Verif Requested. Details: " & question_12_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_13_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q13 Verif Requested. Details: " & question_13_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_14_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q14 Verif Requested. Details: " & question_14_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_15_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q15 Verif Requested. Details: " & question_15_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_16_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q16 Verif Requested. Details: " & question_16_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_17_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q17 Verif Requested. Details: " & question_17_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_18_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q18 Verif Requested. Details: " & question_18_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_19_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q19 Verif Requested. Details: " & question_19_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_20_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q20 Verif Requested. Details: " & question_20_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_21_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q21 Verif Requested. Details: " & question_21_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_22_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q22 Verif Requested. Details: " & question_22_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_23_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q23 Verif Requested. Details: " & question_23_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If question_24_verif_yn = "Requested" Then
					Text 10, y_pos, 500, 10, "Q24 Verif Requested. Details: " & question_24_verif_details
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If

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
				  Text 10, 20, 470, 10, "Note: After you press 'Update' or 'Return to Dialog' the information from the boxes will be added to the list of verification and the boxes will be 'unchecked'."
				  ButtonGroup ButtonPressed
					PushButton 485, 10, 50, 15, "Update", fill_button
			  End If


              ButtonGroup ButtonPressed
                PushButton 540, 10, 60, 15, "Return to Dialog", return_to_dialog_button
              ' Text 10, 340, 580, 50, verifs_needed
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
	Call start_a_blank_case_note
	' Call write_variable_in_CASE_NOTE("CAF Form completed via Phone")
	' Call write_variable_in_CASE_NOTE("Form information taken verbally per COVID Waiver Allowance.")
	' Call write_variable_in_CASE_NOTE("Form information taken on " & caf_form_date)
	' Call write_variable_in_CASE_NOTE("CAF for application date: " & application_date)
	' Call write_variable_in_CASE_NOTE("CAF information saved and will be added to ECF within a few days. Detail can be viewed in 'Assignments Folder'.")
	' Call write_variable_in_CASE_NOTE("---")
	' Call write_variable_in_CASE_NOTE(worker_signature)

	If create_incomplete_note_checkbox = checked then
		CALL write_variable_in_CASE_NOTE("Partial Interview Information from " & interview_date)
	Else
		CALL write_variable_in_CASE_NOTE("~ Interview Completed on " & interview_date & " ~")
	End If
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
	CALL write_variable_in_CASE_NOTE("Completed on " & interview_date & " at " & interview_started_time & " (" & interview_time & " min)")
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
    		If HH_MEMB_ARRAY(client_verification, the_members) <> "Not Needed" Then
    			If HH_MEMB_ARRAY(client_verification, the_members) = "On File" Then
    				If trim(HH_MEMB_ARRAY(client_verification_details, the_members)) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification on file for M" & HH_MEMB_ARRAY(ref_number, the_members) & " - " & HH_MEMB_ARRAY(client_verification_details, the_members))
    				If trim(HH_MEMB_ARRAY(client_verification_details, the_members)) = "" Then CALL write_variable_in_CASE_NOTE("    Verification on file for M" & HH_MEMB_ARRAY(ref_number, the_members) & ".")
    			Else
    				If trim(HH_MEMB_ARRAY(client_verification_details, the_members)) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: of M" & HH_MEMB_ARRAY(ref_number, the_members) & " Information - " & HH_MEMB_ARRAY(client_verification_details, the_members))
    				If trim(HH_MEMB_ARRAY(client_verification_details, the_members)) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: of M" & HH_MEMB_ARRAY(ref_number, the_members) & " Information")
    			End If
    		End If
        End If
	Next
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

	q_1_verbiage = "Q1. Does everyone buy, fix, or eat food together?"
    If question_1_yn <> "" OR trim(question_1_notes) <> "" OR question_1_verif_yn <> "" OR trim(question_1_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE(q_1_verbiage)
    q_1_input = "    CAF Answer - " & question_1_yn
	If question_1_yn <> "" OR trim(question_1_notes) <> "" Then q_1_input = q_1_input & " (Confirmed)"
	If q_1_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_1_input)
	If trim(question_1_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_1_notes)
	If question_1_verif_yn <> "" Then
		If trim(question_1_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_1_verif_yn)
		If trim(question_1_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_1_verif_yn & ": " & question_1_verif_details)
	End If
    If trim(question_1_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_1_interview_notes)

	q_2_verbiage = "Q2. Is anyone (60+) disabled or unable to prepare food?"
    If question_2_yn <> "" OR trim(question_2_notes) <> "" OR question_2_verif_yn <> "" OR trim(question_2_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_2_verbiage)
    q_2_input = "    CAF Answer - " & question_2_yn
	If question_2_yn <> "" OR trim(question_2_notes) <> "" Then q_2_input = q_2_input & " (Confirmed)"
	If q_2_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_2_input)
	If trim(question_2_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_2_notes)
	If question_2_verif_yn <> "" Then
		If trim(question_2_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_2_verif_yn)
		If trim(question_2_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_2_verif_yn & ": " & question_2_verif_details)
	End If
    If trim(question_2_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_2_interview_notes)

    q_3_verbiage = "Q3. Is anyone attending school?"
    If question_3_yn <> "" OR trim(question_3_notes) <> "" OR question_3_verif_yn <> "" OR trim(question_3_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_3_verbiage)
	q_3_input = "    CAF Answer - " & question_3_yn
	If question_3_yn <> "" OR trim(question_3_notes) <> "" Then q_3_input = q_3_input & " (Confirmed)"
	If q_3_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_3_input)
	If trim(question_3_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_3_notes)
	If question_3_verif_yn <> "" Then
		If trim(question_3_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_3_verif_yn)
		If trim(question_3_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_3_verif_yn & ": " & question_3_verif_details)
	End If
    If trim(question_3_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_3_interview_notes)

    q_4_verbiage = "Q4. Is anyone temporarily not living in the home?"
    If question_4_yn <> "" OR trim(question_4_notes) <> "" OR question_4_verif_yn <> "" OR trim(question_4_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_4_verbiage)
	q_4_input = "    CAF Answer - " & question_4_yn
	If question_4_yn <> "" OR trim(question_4_notes) <> "" Then q_4_input = q_4_input & " (Confirmed)"
	If q_4_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_4_input)
	If trim(question_4_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_4_notes)
	If question_4_verif_yn <> "" Then
		If trim(question_4_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_4_verif_yn)
		If trim(question_4_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_4_verif_yn & ": " & question_4_verif_details)
	End If
    If trim(question_4_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_4_interview_notes)

    q_5_verbiage = "Q5. Is anyone blind or does anyone have a limiting illness or disability?"
    If question_5_yn <> "" OR trim(question_5_notes) <> "" OR question_5_verif_yn <> "" OR trim(question_5_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_5_verbiage)
	q_5_input = "    CAF Answer - " & question_5_yn
	If question_5_yn <> "" OR trim(question_5_notes) <> "" Then q_5_input = q_5_input & " (Confirmed)"
	If q_5_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_5_input)
	If trim(question_5_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_5_notes)
	If question_5_verif_yn <> "" Then
		If trim(question_5_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_5_verif_yn)
		If trim(question_5_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_5_verif_yn & ": " & question_5_verif_details)
	End If
    If trim(question_5_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_5_interview_notes)

    q_6_verbiage = "Q6. Is anyone unable to work?"
    If question_6_yn <> "" OR trim(question_6_notes) <> "" OR question_6_verif_yn <> "" OR trim(question_6_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_6_verbiage)
	q_6_input = "    CAF Answer - " & question_6_yn
	If question_6_yn <> "" OR trim(question_6_notes) <> "" Then q_6_input = q_6_input & " (Confirmed)"
	If q_6_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_6_input)
	If trim(question_6_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_6_notes)
	If question_6_verif_yn <> "" Then
		If trim(question_6_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_6_verif_yn)
		If trim(question_6_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_6_verif_yn & ": " & question_6_verif_details)
	End If
    If trim(question_6_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_6_interview_notes)

    q_7_verbiage = "Q7. Has anyone stopped, quit or refused employment in the past 60 days?"
    If question_7_yn <> "" OR trim(question_7_notes) <> "" OR question_7_verif_yn <> "" OR trim(question_7_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_7_verbiage)
	q_7_input = "    CAF Answer - " & question_7_yn
	If question_7_yn <> "" OR trim(question_7_notes) <> "" Then q_7_input = q_7_input & " (Confirmed)"
	If q_7_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_7_input)
	If trim(question_7_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_7_notes)
	If question_7_verif_yn <> "" Then
		If trim(question_7_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_7_verif_yn)
		If trim(question_7_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_7_verif_yn & ": " & question_7_verif_details)
	End If
    If trim(question_7_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_7_interview_notes)

    q_8_verbiage = "Q8. Has anyone had a job OR been self-employed in the past 12 months?"
    If question_8_yn <> "" OR trim(question_8_notes) <> "" OR question_8a_yn <> "" OR question_8_verif_yn <> "" OR trim(question_8_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_8_verbiage)
	q_8_input = "    CAF Answer - " & question_8_yn
    If question_8_yn <> "" OR trim(question_8_notes) <> "" Then q_8_input = q_8_input & " (Confirmed)"
    If q_8_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_8_input)
    q_8a_verbiage = "    Q8a.In the past 36 months? (SNAP ONLY)"
	If question_8a_yn <> "" Then
        Call write_variable_in_CASE_NOTE(q_8a_verbiage)
        Call write_variable_in_CASE_NOTE("        CAF Answer - " & question_8a_yn)
    End If
	If trim(question_8_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_8_notes)
	If question_8_verif_yn <> "" Then
		If trim(question_8_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_8_verif_yn)
		If trim(question_8_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_8_verif_yn & ": " & question_8_verif_details)
	End If
    If trim(question_8_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_8_interview_notes)


    q_9_verbiage = "Q9. Does anyone have a job?"
    If question_9_yn <> "" OR trim(question_9_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_9_verbiage)
    ' If question_9_yn <> "" OR trim(question_9_notes) <> "" OR question__verif_yn <> "" OR trim(question__interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_9_verbiage)
	q_9_input = "    CAF Answer - " & question_9_yn
	If question_9_yn <> "" OR trim(question_9_notes) <> "" Then q_9_input = q_9_input & " (Confirmed)"
	If q_9_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_9_input)
	for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
		If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
			CALL write_variable_in_CASE_NOTE("    Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & " for " & JOBS_ARRAY(jobs_employee_name, each_job) & " monthly earnings $" & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job))
			If JOBS_ARRAY(verif_yn, each_job) <> "" Then
				If trim(JOBS_ARRAY(verif_details, each_job)) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & JOBS_ARRAY(verif_yn, each_job))
				If trim(JOBS_ARRAY(verif_details, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & JOBS_ARRAY(verif_yn, each_job) & ": " & JOBS_ARRAY(verif_details, each_job))
			End If
			If trim(JOBS_ARRAY(jobs_notes, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer: " & JOBS_ARRAY(jobs_notes, each_job))
			If trim(JOBS_ARRAY(jobs_intv_notes, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & JOBS_ARRAY(jobs_intv_notes, each_job))
		End If
	next

    q_10_verbiage = "Q10.Is anyone self-employed?"
    If question_10_yn <> "" OR trim(question_10_notes) <> "" OR question_10_verif_yn <> "" OR trim(question_10_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_10_verbiage)
	q_10_input = "    CAF Answer - " & question_10_yn
	If trim(question_10_monthly_earnings) <> "" Then q_10_input = q_10_input & " Gross Monthly Earnings: " & question_10_monthly_earnings
	If question_10_yn <> "" OR trim(question_10_notes) <> "" Then q_10_input = q_10_input & " (Confirmed)"
	If q_10_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_10_input)
	If trim(question_10_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_10_notes)
	If question_10_verif_yn <> "" Then
		If trim(question_10_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_10_verif_yn)
		If trim(question_10_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_10_verif_yn & ": " & question_10_verif_details)
	End If
    If trim(question_10_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_10_interview_notes)

    q_11_verbiage = "Q11.Do you expect any changes in income, expenses, or work hours?"
    If question_11_yn <> "" OR trim(question_11_notes) <> "" OR question_11_verif_yn <> "" OR trim(question_11_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_11_verbiage)
	q_11_input = "    CAF Answer - " & question_11_yn
	If question_11_yn <> "" OR trim(question_11_notes) <> "" Then q_11_input = q_11_input & " (Confirmed)"
	If q_11_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_11_input)
	If trim(question_11_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_11_notes)
	If question_11_verif_yn <> "" Then
		If trim(question_11_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_11_verif_yn)
		If trim(question_11_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_11_verif_yn & ": " & question_11_verif_details)
	End If
    If trim(question_11_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_11_interview_notes)

	If trim(pwe_selection) <> "" AND pwe_selection <> "Select or Type" Then CALL write_variable_in_CASE_NOTE("PWE: " & pwe_selection)

    q_12_verbiage = "Q12.Does anyone have any unearned income?"
	If q_12_totally_blank = False Then
        Call write_variable_in_CASE_NOTE(q_12_verbiage)
		CALL write_variable_in_CASE_NOTE("    CAF Answer:")

		question_12_rsdi_yn = left(question_12_rsdi_yn & "   ", 5)
		If trim(question_12_rsdi_amt) <> "" Then question_12_rsdi_amt = left("$" & question_12_rsdi_amt & ".00       ", 8)
		question_12_ssi_yn = left(question_12_ssi_yn & "   ", 5)
		If trim(question_12_ssi_amt) <> "" Then question_12_ssi_amt = left("$" & question_12_ssi_amt & ".00       ", 8)
		question_12_va_yn = left(question_12_va_yn & "   ", 5)
		If trim(question_12_va_amt) <> "" Then question_12_va_amt = left("$" & question_12_va_amt & ".00       ", 8)
		question_12_ui_yn = left(question_12_ui_yn & "   ", 5)
		If trim(question_12_ui_amt) <> "" Then question_12_ui_amt = left("$" & question_12_ui_amt & ".00       ", 8)
		question_12_wc_yn = left(question_12_wc_yn & "   ", 5)
		If trim(question_12_wc_amt) <> "" Then question_12_wc_amt = left("$" & question_12_wc_amt & ".00       ", 8)
		question_12_ret_yn = left(question_12_ret_yn & "   ", 5)
		If trim(question_12_ret_amt) <> "" Then question_12_ret_amt = left("$" & question_12_ret_amt & ".00       ", 8)
		question_12_trib_yn = left(question_12_trib_yn & "   ", 5)
		If trim(question_12_trib_amt) <> "" Then question_12_trib_amt = left("$" & question_12_trib_amt & ".00       ", 8)
		question_12_cs_yn = left(question_12_cs_yn & "   ", 5)
		If trim(question_12_cs_amt) <> "" Then question_12_cs_amt = left("$" & question_12_cs_amt & ".00       ", 8)
		question_12_other_yn = left(question_12_other_yn & "   ", 5)
		If trim(question_12_other_amt) <> "" Then question_12_other_amt = left("$" & question_12_other_amt & ".00       ", 8)


		CALL write_variable_in_CASE_NOTE("    RSDI - " & question_12_rsdi_yn & " " & question_12_rsdi_amt & "   UI - " & question_12_ui_yn & " " & question_12_ui_amt & " Tribal - " & question_12_trib_yn & " " & question_12_trib_amt)
		CALL write_variable_in_CASE_NOTE("     SSI - " & question_12_ssi_yn & " " & question_12_ssi_amt & "   WC - " & question_12_wc_yn & " " & question_12_wc_amt & "   CSES - " & question_12_cs_yn & " " & question_12_cs_amt)
		CALL write_variable_in_CASE_NOTE("      VA - " & question_12_va_yn & " " & question_12_va_amt & "  Ret - " & question_12_ret_yn & " " & question_12_ret_amt & "  Other - " & question_12_other_yn & " " & question_12_other_amt)
		If trim(question_12_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_12_notes)
	End If
	If question_12_verif_yn <> "" Then
		If trim(question_12_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_12_verif_yn)
		If trim(question_12_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_12_verif_yn & ": " & question_12_verif_details)
	End If
    If trim(question_12_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_12_interview_notes)

    q_13_verbiage = "Q13.Does anyone receive financial aid for attending school?"
    If question_13_yn <> "" OR trim(question_13_notes) <> "" OR question_13_verif_yn <> "" OR trim(question_13_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_13_verbiage)
	q_13_input = "    CAF Answer - " & question_13_yn
	If question_13_yn <> "" OR trim(question_13_notes) <> "" Then q_13_input = q_13_input & " (Confirmed)"
	If q_13_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_13_input)
	If trim(question_13_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_13_notes)
	If question_13_verif_yn <> "" Then
		If trim(question_13_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_13_verif_yn)
		If trim(question_13_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_13_verif_yn & ": " & question_13_verif_details)
	End If
    If trim(question_13_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_13_interview_notes)

    q_14_verbiage = "Q14.Are there any of the following housing expenses?"
	If q_14_totally_blank = False Then
        Call write_variable_in_CASE_NOTE(q_14_verbiage)
		CALL write_variable_in_CASE_NOTE("    CAF Answer:")

		question_14_rent_yn = left(question_14_rent_yn & "   ", 5)
		question_14_subsidy_yn = left(question_14_subsidy_yn & "   ", 5)
		question_14_mortgage_yn = left(question_14_mortgage_yn & "   ", 5)
		' question_14_taxes_yn = left(question_14_taxes_yn & "   ", 5)
		question_14_association_yn = left(question_14_association_yn & "   ", 5)
		' question_14_insurance_yn = left(question_14_insurance_yn & "   ", 5)
		question_14_room_yn = left(question_14_room_yn & "   ", 5)

		' CALL write_variable_in_CASE_NOTE("       Rent - " & question_14_rent_yn        & " Rental Subsidy - " & question_14_subsidy_yn & "  Mortgage - " & question_14_mortgage_yn & " Taxes - " & question_14_taxes_yn)
		' CALL write_variable_in_CASE_NOTE(" Assoc Fees - " & question_14_association_yn & "     Room/Board - " & question_14_room_yn    & " Insurance - " & question_14_insurance_yn)
		CALL write_variable_in_CASE_NOTE("       Rent - " & question_14_rent_yn &  " Rental Subsidy - " & question_14_subsidy_yn & "  Mortgage - " & question_14_mortgage_yn & "    Taxes - " & question_14_taxes_yn)
		CALL write_variable_in_CASE_NOTE("                        Assoc Fees - " & question_14_association_yn & "Room/Board - " & question_14_room_yn    & "Insurance - " & question_14_insurance_yn)
        If trim(question_14_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_14_notes)
	End If
	If question_14_verif_yn <> "" Then
		If trim(question_14_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_14_verif_yn)
		If trim(question_14_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_14_verif_yn & ": " & question_14_verif_details)
	End If
    If trim(question_14_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_14_interview_notes)
	If disc_rent_amounts = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    ANSWER MAY NOT MATCH CAF PG 1 INFORMATION")
		CALL write_variable_in_CASE_NOTE("    Resolution: " & disc_rent_amounts_confirmation)
	End If
	If disc_yes_phone_no_expense = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    PHONE NUMBER LISTED BUT NO PHONE EXPENSE")
		CALL write_variable_in_CASE_NOTE("    Resolution: " & disc_yes_phone_no_expense_confirmation)
	End If
	If disc_no_phone_yes_expense = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    NO PHONE NUMBER LISTED BUT EXPENSE EXISTS")
		CALL write_variable_in_CASE_NOTE("    Resolution: " & disc_no_phone_yes_expense_confirmation)
	End If

    q_15_verbiage = "Q15.Are there any of the following utility expenses?"
	If q_15_totally_blank = False Then
        Call write_variable_in_CASE_NOTE(q_15_verbiage)
		CALL write_variable_in_CASE_NOTE("    CAF Answer:")

		question_15_heat_ac_yn = left(question_15_heat_ac_yn & "   ", 5)
		question_15_electricity_yn = left(question_15_electricity_yn & "   ", 5)
		' question_15_cooking_fuel_yn = left(question_15_cooking_fuel_yn & "   ", 5)
		question_15_water_and_sewer_yn = left(question_15_water_and_sewer_yn & "   ", 5)
		question_15_garbage_yn = left(question_15_garbage_yn & "   ", 5)
		' question_15_phone_yn = left(question_15_phone_yn & "   ", 5)
		' question_15_liheap_yn = left(question_15_liheap_yn & "   ", 5)

		CALL write_variable_in_CASE_NOTE("        Heat/AC - " & question_15_heat_ac_yn & " Electric - " & question_15_electricity_yn & " Cooking Fuel - " & question_15_cooking_fuel_yn)
		CALL write_variable_in_CASE_NOTE("    Water/Sewer - " & question_15_water_and_sewer_yn & "  Garbage - " & question_15_garbage_yn & "        Phone - " & question_15_phone_yn)
        CALL write_variable_in_CASE_NOTE("    LIHEAP/Energy Assistance in past 12 months - " & question_15_liheap_yn)
		If trim(question_15_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_15_notes)
	End If
	If question_15_verif_yn <> "" Then
		If trim(question_15_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_15_verif_yn)
		If trim(question_15_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_15_verif_yn & ": " & question_15_verif_details)
	End If
    If trim(question_15_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_15_interview_notes)
    If trim(question_15_phone_details) <> "" AND question_15_phone_details <> "Select or Type" Then CALL write_variable_in_CASE_NOTE("    PHONE DETAILS: " & question_15_phone_details)

	If disc_utility_amounts = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    ANSWER MAY NOT MATCH CAF PG 1 INFORMATION")
		CALL write_variable_in_CASE_NOTE("    Resolution: " & disc_utility_amounts_confirmation)
	End If

    q_16_verbiage = "Q16.Does anyone have costs for childcare?"
    If question_16_yn <> "" OR trim(question_16_notes) <> "" OR question_16_verif_yn <> "" OR trim(question_16_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_16_verbiage)
	q_16_input = "    CAF Answer - " & question_16_yn
	If question_16_yn <> "" OR trim(question_16_notes) <> "" Then q_16_input = q_16_input & " (Confirmed)"
	If q_16_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_16_input)
	If trim(question_16_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_16_notes)
	If question_16_verif_yn <> "" Then
		If trim(question_16_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_16_verif_yn)
		If trim(question_16_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_16_verif_yn & ": " & question_16_verif_details)
	End If
    If trim(question_16_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_16_interview_notes)

    q_17_verbiage = "Q17.Does anyone have costs for adult care?"
    If question_17_yn <> "" OR trim(question_17_notes) <> "" OR question_17_verif_yn <> "" OR trim(question_17_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_17_verbiage)
	q_17_input = "    CAF Answer - " & question_17_yn
	If question_17_yn <> "" OR trim(question_17_notes) <> "" Then q_17_input = q_17_input & " (Confirmed)"
	If q_17_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_17_input)
	If trim(question_17_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_17_notes)
	If question_17_verif_yn <> "" Then
		If trim(question_17_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_17_verif_yn)
		If trim(question_17_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_17_verif_yn & ": " & question_17_verif_details)
	End If
    If trim(question_17_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_17_interview_notes)

    q_18_verbiage = "Q18.Does anyone pay support to someone outside of the home?"
    If question_18_yn <> "" OR trim(question_18_notes) <> "" OR question_18_verif_yn <> "" OR trim(question_18_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_18_verbiage)
	q_18_input = "    CAF Answer - " & question_18_yn
	If question_18_yn <> "" OR trim(question_18_notes) <> "" Then q_18_input = q_18_input & " (Confirmed)"
	If q_18_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_18_input)
	If trim(question_18_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_18_notes)
	If question_18_verif_yn <> "" Then
		If trim(question_18_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_18_verif_yn)
		If trim(question_18_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_18_verif_yn & ": " & question_18_verif_details)
	End If
    If trim(question_18_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_18_interview_notes)

    q_19_verbiage = "Q19.Does anyone (disabled or 60+) have medical expenses? (SNAP ONLY)"
    If question_19_yn <> "" OR trim(question_19_notes) <> "" OR question_19_verif_yn <> "" OR trim(question_19_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_19_verbiage)
	q_19_input = "    CAF Answer - " & question_19_yn
	If question_19_yn <> "" OR trim(question_19_notes) <> "" Then q_19_input = q_19_input & " (Confirmed)"
	If q_19_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_19_input)
	If trim(question_19_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_19_notes)
	If question_19_verif_yn <> "" Then
		If trim(question_19_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_19_verif_yn)
		If trim(question_19_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_19_verif_yn & ": " & question_19_verif_details)
	End If
    If trim(question_19_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_19_interview_notes)

    q_20_verbiage = "Q20.Does anyone own or is anyone buying any of the following:"
	If q_20_totally_blank = False Then
        Call write_variable_in_CASE_NOTE(q_20_verbiage)
		CALL write_variable_in_CASE_NOTE("    CAF Answer:")

		question_20_cash_yn = left(question_20_cash_yn & "   ", 5)
		question_20_acct_yn = left(question_20_acct_yn & "   ", 5)
		question_20_secu_yn = left(question_20_secu_yn & "   ", 5)
		question_20_cars_yn = left(question_20_cars_yn & "   ", 5)


		CALL write_variable_in_CASE_NOTE("      Cash - " & question_20_cash_yn & " Bank Accounts - " & question_20_acct_yn)
		CALL write_variable_in_CASE_NOTE("    Stocks - " & question_20_secu_yn & "      Vehicles - " & question_20_cars_yn)
		If trim(question_20_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_20_notes)
	End If
	If question_20_verif_yn <> "" Then
		If trim(question_20_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_20_verif_yn)
		If trim(question_20_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_20_verif_yn & ": " & question_20_verif_details)
	End If
    If trim(question_20_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_20_interview_notes)

    q_21_verbiage = "Q21.Has anyone sold/given away/traded assets in the past 12 mos?(CASH ONLY)"
    If question_21_yn <> "" OR trim(question_21_notes) <> "" OR question_21_verif_yn <> "" OR trim(question_21_interview_notes) <> "" Then Call write_variable_in_CASE_NOTE(q_21_verbiage)
	q_21_input = "    CAF Answer - " & question_21_yn
	If question_21_yn <> "" OR trim(question_21_notes) <> "" Then q_21_input = q_21_input & " (Confirmed)"
	If q_21_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_21_input)
	If trim(question_21_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_21_notes)
	If question_21_verif_yn <> "" Then
		If trim(question_21_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_21_verif_yn)
		If trim(question_21_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_21_verif_yn & ": " & question_21_verif_details)
	End If
    If trim(question_21_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_21_interview_notes)

    q_22_verbiage = "Q22.Did anyone move in/out in the past 12 months? (REVW ONLY)"
    If question_22_yn <> "" OR trim(question_22_notes) <> "" OR question_22_verif_yn <> "" OR trim(question_22_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE(q_22_verbiage)
    q_22_input = "    CAF Answer - " & question_22_yn
	If question_22_yn <> "" OR trim(question_22_notes) <> "" Then q_22_input = q_22_input & " (Confirmed)"
	If q_22_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_22_input)
    If trim(question_22_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_22_notes)
	If question_22_verif_yn <> "" Then
		If trim(question_22_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_22_verif_yn)
		If trim(question_22_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_22_verif_yn & ": " & question_22_verif_details)
	End If
    If trim(question_22_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_22_interview_notes)

    q_23_verbiage = "Q23.Are both parents of children under 19 living in the home?"
    If question_23_yn <> "" OR trim(question_23_notes) <> "" OR question_23_verif_yn <> "" OR trim(question_23_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE(q_23_verbiage)
	q_23_input = "    CAF Answer - " & question_23_yn
	If question_23_yn <> "" OR trim(question_23_notes) <> "" Then q_23_input = q_23_input & " (Confirmed)"
	If q_23_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_23_input)
	If trim(question_23_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_23_notes)
	If question_23_verif_yn <> "" Then
		If trim(question_23_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_23_verif_yn)
		If trim(question_23_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_23_verif_yn & ": " & question_23_verif_details)
	End If
    If trim(question_23_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_23_interview_notes)

    q_24_verbiage = "Q24.Does anyone have any of the following expenses? (MSA ONLY)"
	If q_24_totally_blank = False Then
        Call write_variable_in_CASE_NOTE(q_24_verbiage)
		CALL write_variable_in_CASE_NOTE("    CAF Answer:")
		question_24_rep_payee_yn = left(question_24_rep_payee_yn & "   ", 5)
		question_24_guardian_fees_yn = left(question_24_guardian_fees_yn & "   ", 5)
		question_24_special_diet_yn = left(question_24_special_diet_yn & "   ", 5)
		question_24_high_housing_yn = left(question_24_high_housing_yn & "   ", 5)

		CALL write_variable_in_CASE_NOTE("    REP Payee Fees - " & question_24_rep_payee_yn    & "         Guard Fees - " & question_24_guardian_fees_yn)
		CALL write_variable_in_CASE_NOTE("      Special Diet - " & question_24_special_diet_yn & " High Housing Costs - " & question_24_high_housing_yn)
		If trim(question_24_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & question_24_notes)
	End If
	If question_24_verif_yn <> "" Then
		If trim(question_24_verif_details) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_24_verif_yn)
		If trim(question_24_verif_details) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & question_24_verif_yn & ": " & question_24_verif_details)
	End If
    If trim(question_24_interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & question_24_interview_notes)

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

	IF create_verif_note = True Then Call write_variable_in_CASE_NOTE("** VERIFICATIONS REQUESTED - See previous case note for detail")
	IF create_verif_note = False Then Call write_variable_in_CASE_NOTE("No verifications were indicated at this time.")

    If IsArray(note_detail_array) = True Then
    	first_resource = True
    	For each note_line in note_detail_array
    		IF note_line <> "" Then
    			If first_resource = True Then
    				call write_variable_in_CASE_NOTE("Additional resource information given to resident")
    				first_resource = False
    			End If
    			Call write_variable_in_CASE_NOTE(note_line)
    		End If
    	Next
    End If

	If qual_questions_yes = FALSE Then Call write_variable_in_CASE_NOTE("* All CAF Qualifying Questions answered 'No'.")

	forms_reviewed = ""
	If left(confirm_resp_read, 4) = "YES!" AND left(confirm_rights_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -4163"
	If left(confirm_ebt_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -EBT Info"
	If left(confirm_ebt_how_to_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -3315A"
	If left(confirm_npp_info_read, 4) = "YES!" AND left(confirm_npp_rights_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -3979"
	If left(confirm_ievs_info_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -2759"
	If left(confirm_appeal_rights_read, 4) = "YES!" AND left(confirm_civil_rights_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -3353"
	If left(confirm_cover_letter_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -Hennepin County Information "
	If left(confirm_program_information_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -2920"
	If left(confirm_DV_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -3477"
	If left(confirm_disa_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -4133"
	If left(confirm_mfip_forms_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -2647 -2929 -3323"
	If left(confirm_mfip_cs_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -3393 -3163B -2338 -5561"
	If left(confirm_minor_mfip_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -2961 -2887 -3238"
	If left(confirm_snap_forms_read, 4) = "YES!" Then forms_reviewed = forms_reviewed & " -2625 -2707 -7635"
	If left(forms_reviewed, 2) = " -" Then forms_reviewed = right(forms_reviewed, len(forms_reviewed)-2)
	Call write_bullet_and_variable_in_CASE_NOTE("Reviewed DHS Forms", forms_reviewed)
	If left(confirm_snap_forms_read, 4) = "YES!" Then
		Call write_variable_in_CASE_NOTE("SNAP Reporting discussed. Case appears to be a " & snap_reporting_type & " reporter.")
        Call write_variable_in_CASE_NOTE("     Next review month of " & next_revw_month)
		Call write_variable_in_CASE_NOTE("     This may change dependent on info received up until SNAP approval.")
	End If


	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

end function

function create_verifs_needed_list(verifs_selected, verifs_needed)

	verifs_needed = verifs_selected
	If right(verifs_needed, 1) = ";" Then verifs_needed = left(verifs_needed, len(verifs_needed) - 1)
	If left(verifs_needed, 1) = ";" Then verifs_needed = right(verifs_needed, len(verifs_needed) - 1)

	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
        If HH_MEMB_ARRAY(ignore_person, the_members) = False Then
            If HH_MEMB_ARRAY(client_verification, the_members) = "Requested" Then
    			verifs_needed = verifs_needed & "; MEMB " & HH_MEMB_ARRAY(ref_number, the_members) & "-" & HH_MEMB_ARRAY(full_name_const, the_members) & " Information. "
    			If trim(HH_MEMB_ARRAY(client_verification_details, the_members)) <> "" Then verifs_needed = verifs_needed & " - " & HH_MEMB_ARRAY(client_verification_details, the_members)
    		End If
        End If
	Next
	If question_1_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q1 Information (P&P Together)"
		If trim(question_1_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_1_verif_details
	End If
	If question_2_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q2 Information (Ages/DISA unable to buy food)"
		If trim(question_2_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_2_verif_details
	End If

	If question_3_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q3 Information (Attending School)"
		If trim(question_3_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_3_verif_details
	End If

	If question_4_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q4 Information (Temp out of Home)"
		If trim(question_4_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_4_verif_details
	End If

	If question_5_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q5 Information (DISA)"
		If trim(question_5_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_5_verif_details
	End If

	If question_6_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q6 Information (Unable to Work)"
		If trim(question_6_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_6_verif_details
	End If

	If question_7_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q7 Information (Job end/reduce in past 60 Days)"
		If trim(question_7_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_7_verif_details
	End If

	If question_8_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q8 Information (Employed in past 12 Months)"
		If trim(question_8_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_8_verif_details
	End If

	For each_job = 0 to UBound(JOBS_ARRAY, 2)
		If JOBS_ARRAY(verif_yn, each_job) = "Requested" Then
			verifs_needed = verifs_needed & "; CAF Q9 Information (Job) - " & JOBS_ARRAY(employer_name, each_job)
			If trim(JOBS_ARRAY(verif_details, each_job)) <> "" Then verifs_needed = verifs_needed & " - " & JOBS_ARRAY(verif_details, each_job)
		End If

	Next
	If question_10_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q10 Information (Self Employed)"
		If trim(question_10_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_10_verif_details
	End If

	If question_11_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q11 Information (Income Changes)"
		If trim(question_11_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_11_verif_details
	End If

	If question_12_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q12 Information (UNEA Income)"
		If trim(question_12_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_12_verif_details
	End If

	If question_13_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q13 Information (School Financial Aid)"
		If trim(question_13_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_13_verif_details
	End If

	If question_14_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q14 Information (Housing Expense)"
		If trim(question_14_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_14_verif_details
	End If

	If question_15_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q15 Information (Utilities Expense)"
		If trim(question_15_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_15_verif_details
	End If

	If question_16_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q16 Information (Child Care Expense)"
		If trim(question_16_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_16_verif_details
	End If

	If question_17_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q17 Information (DISA Adult Care Expense)"
		If trim(question_17_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_17_verif_details
	End If

	If question_18_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q18 Information (Child Support Expense)"
		If trim(question_18_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_18_verif_details
	End If

	If question_19_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q19 Information (Medical Expenses)"
		If trim(question_19_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_19_verif_details
	End If

	If question_20_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q20 Information (Assets)"
		If trim(question_20_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_20_verif_details
	End If

	If question_21_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q21 Information (Asset Trade)"
		If trim(question_21_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_21_verif_details
	End If

	If question_22_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q22 Information (Anyone Move In or Out)"
		If trim(question_22_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_22_verif_details
	End If

	If question_23_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q23 Information (Both Parents in Home)"
		If trim(question_23_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_23_verif_details
	End If

	If question_24_verif_yn = "Requested" Then
		verifs_needed = verifs_needed & "; CAF Q24 Information (MSA Expenses)"
		If trim(question_24_verif_details) <> "" Then verifs_needed = verifs_needed & " - " & question_24_verif_details
	End If

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
		BeginDialog Dialog1, 0, 0, 296, 160, "Determination of Assets in Month of Application"
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

			dlg_len = 45 + jobs_grp_len + busi_grp_len + unea_grp_len

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, dlg_len, "Determination of Income in Month of Application"
			  ' Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  ButtonGroup ButtonPressed
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
			dlg_len = 55 + cash_grp_len + acct_grp_len

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 351, dlg_len, "Determination of Assets in Month of Application"
			  Text 10, 10, 205, 10, "Are there any Liquid Assets available to the household?"
			  GroupBox 10, 25, 220, cash_grp_len, "Cash"
			  If cash_amount_yn = "Yes" Then
				  Text 20, 40, 155, 10, "This household HAS Cash Savings."
				  Text 20, 55, 150, 10, "How much in total does the household have?"
				  EditBox 175, 50, 45, 15, cash_amount
				  y_pos = 80
			  Else
				  Text 20, 40, 155, 10, "This household does NOT have Cash."
				  y_pos = 60
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
	Do
		prvt_err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 196, 140, "Determination of Housing Cost in Month of Application"
		  EditBox 45, 35, 50, 15, rent_amount
		  EditBox 45, 55, 50, 15, lot_rent_amount
		  EditBox 45, 75, 50, 15, mortgage_amount
		  EditBox 45, 95, 50, 15, insurance_amount
		  EditBox 140, 35, 50, 15, tax_amount
		  EditBox 140, 55, 50, 15, room_amount
		  EditBox 140, 75, 50, 15, garage_amount
		  EditBox 140, 95, 50, 15, subsidy_amount
		  ButtonGroup ButtonPressed
		    PushButton 140, 120, 50, 15, "Return", return_btn
		  Text 10, 15, 165, 10, "Enter the total Shelter Expense for the Houshold."
		  Text 25, 40, 20, 10, "Rent:"
		  Text 10, 60, 35, 10, " Lot Rent:"
		  Text 10, 80, 35, 10, "Mortgage:"
		  Text 10, 100, 35, 10, "Insurance:"
		  Text 115, 40, 25, 10, "Taxes:"
		  Text 115, 60, 25, 10, "Room:"
		  Text 110, 80, 30, 10, "Garage:"
		  Text 105, 100, 35, 10, "  Subsidy:"
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
	' If IsNumeric(subsidy_amount) = True Then determined_shel = determined_shel + subsidy_amount

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

	Do
		current_utilities = all_utilities

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 246, 175, "Determination of Utilities in Month of Application"
		  CheckBox 30, 45, 50, 10, "Heat", heat_checkbox
		  CheckBox 30, 60, 65, 10, "Air Conditioning", ac_checkbox
		  CheckBox 30, 75, 50, 10, "Electric", electric_checkbox
		  CheckBox 30, 90, 50, 10, "Phone", phone_checkbox
		  CheckBox 30, 105, 50, 10, "NONE", none_checkbox
		  ButtonGroup ButtonPressed
		    PushButton 170, 105, 65, 15, "Calculate", calculate_btn
		    PushButton 170, 155, 65, 15, "Return", return_btn
		  Text 10, 10, 235, 10, "Check the boxes for each utility the household is responsible to pay:"
		  GroupBox 15, 30, 225, 95, "Utilities"
		  Text 150, 45, 50, 10, "$ " & determined_utilities
		  Text 150, 60, 35, 35, all_utilities
		  Text 15, 135, 225, 20, "Remember, this expense could be shared, they are still considered responsible to pay and we count the WHOLE standard."
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

		add_msg = "Denials can be coded in REPT/PND2 if they are for a resident 'Withdraw' of their request. Otherise, since the interview should be done at this point, denials should be processed in STAT."
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
				' If DateDiff("d", mn_elig_begin_date, CAF_datestamp) > 0 Then
				' 	mn_elig_begin_date = CAF_datestamp
				' 	expedited_package = original_expedited_package
				' Else
				' 	MN_elig_month = DatePart("m", mn_elig_begin_date)
				' 	MN_elig_month = right("0"&MN_elig_month, 2)
				' 	MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
				' 	expedited_package = MN_elig_month & "/" & MN_elig_year
				' End If
			End If
		ElseIf IsDate(other_state_reported_benefit_end_date) = True Then
			If DateDiff("d", day_30_from_application, other_state_reported_benefit_end_date) >= 0 Then
				action_due_to_out_of_state_benefits = "DENY"
			Else
				action_due_to_out_of_state_benefits = "APPROVE"
				mn_elig_begin_date = DateAdd("d", 1, other_state_reported_benefit_end_date)
				' If DateDiff("d", mn_elig_begin_date, CAF_datestamp) > 0 Then
				' 	mn_elig_begin_date = CAF_datestamp
				' 	expedited_package = original_expedited_package
				' Else
				' 	MN_elig_month = DatePart("m", mn_elig_begin_date)
				' 	MN_elig_month = right("0"&MN_elig_month, 2)
				' 	MN_elig_year = right(DatePart("yyyy", mn_elig_begin_date), 2)
				' 	expedited_package = MN_elig_month & "/" & MN_elig_year
				' End If
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
	call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", True)
	' call create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", email_subject, email_body, "", False)
	' create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
end function
'---------------------------------------------------------------------------------------------------------------------------


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

const jobs_employee_name 			= 0
const jobs_hourly_wage 				= 1
const jobs_gross_monthly_earnings	= 2
const jobs_employer_name 			= 3
const jobs_edit_btn					= 4
const jobs_intv_notes				= 5
const verif_yn						= 6
const verif_details					= 7
const jobs_notes 					= 8

Const end_of_doc = 6			'This is for word document ennumeration

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
' worker_name = user_ID_for_validation
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
Dim JOBS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)
ReDim JOBS_ARRAY(jobs_notes, 0)

const account_type_const	= 0
const account_owner_const	= 1
const bank_name_const		= 2
const account_amount_const	= 3
const account_notes_const 	= 4

Dim EXP_ACCT_ARRAY
ReDim EXP_ACCT_ARRAY(account_notes_const, 0)

const jobs_employee_const 	= 0
const jobs_employer_const	= 1
const jobs_wage_const		= 2
const jobs_hours_const		= 3
const jobs_frequency_const 	= 4
const jobs_monthly_pay_const= 5
const jobs_notes_const 		= 6

Dim EXP_JOBS_ARRAY
ReDim EXP_JOBS_ARRAY(jobs_notes_const, 0)

const busi_owner_const 				= 0
const busi_info_const 				= 1
const busi_monthly_earnings_const	= 2
const busi_annual_earnings_const	= 3
const busi_notes_const 				= 4

Dim EXP_BUSI_ARRAY
ReDim EXP_BUSI_ARRAY(busi_notes_const, 0)

const unea_owner_const 				= 0
const unea_info_const 				= 1
const unea_monthly_earnings_const	= 2
const unea_weekly_earnings_const	= 3
const unea_notes_const 				= 4

Dim EXP_UNEA_ARRAY
ReDim EXP_UNEA_ARRAY(unea_notes_const, 0)

Call remove_dash_from_droplist(state_list)
'These are all the definitions for droplists

memb_panel_relationship_list = "Select One..."
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"01 Applicant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"02 Spouse"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"03 Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"04 Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"05 Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"06 Step Sibling"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"08 Step Child"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"09 Step Parent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"10 Aunt"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"11 Uncle"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"12 Niece"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"13 Nephew"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"14 Cousin"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"15 Grandparent"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"16 Grandchild"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"17 Other Relative"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"18 Legal Guardian"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"24 Not Related"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"25 Live-In Attendant"
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"27 Unknown"

marital_status_list = "Select One..."
marital_status_list = marital_status_list+chr(9)+"N  Never Married"
marital_status_list = marital_status_list+chr(9)+"M  Married Living With Spouse"
marital_status_list = marital_status_list+chr(9)+"S  Married Living Apart (Sep)"
marital_status_list = marital_status_list+chr(9)+"L  Legally Sep"
marital_status_list = marital_status_list+chr(9)+"D  Divorced"
marital_status_list = marital_status_list+chr(9)+"W  Widowed"

id_droplist_info = "BC - Birth Certificate"
id_droplist_info = id_droplist_info+chr(9)+"RE - Religious Record"
id_droplist_info = id_droplist_info+chr(9)+"DL - Drivers License/ST ID"
id_droplist_info = id_droplist_info+chr(9)+"DV - Divorce Decree"
id_droplist_info = id_droplist_info+chr(9)+"AL - Alien Card"
id_droplist_info = id_droplist_info+chr(9)+"AD - Arrival//Depart"
id_droplist_info = id_droplist_info+chr(9)+"DR - Doctor Stmt"
id_droplist_info = id_droplist_info+chr(9)+"PV - Passport/Visa"
id_droplist_info = id_droplist_info+chr(9)+"OT - Other Document"
id_droplist_info = id_droplist_info+chr(9)+"NO - No Ver Prvd"
id_droplist_info = id_droplist_info+chr(9)+"Found in SOLQ/SMI"
id_droplist_info = id_droplist_info+chr(9)+"Requested"

question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"

Set wshshell = CreateObject("WScript.Shell")						'creating the wscript method to interact with the system
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"	'defining the my documents folder for use in saving script details/variables between script runs
If user_ID_for_validation = "ERHO003" Then user_c_drive_docs_folder = "C:\Users\" & lcase(windows_user_ID) & "\Documents\"

'Dimming all the variables because they are defined and set within functions
Dim who_are_we_completing_the_interview_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, exp_q_4_utilities_this_month, caf_exp_pay_heat_checkbox, caf_exp_pay_ac_checkbox, caf_exp_pay_electricity_checkbox, caf_exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_one_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county, CAF_datestamp, all_the_clients, err_msg, interpreter_information, interpreter_language, arep_interview_id_information, non_applicant_interview_info
Dim intv_app_month_income, intv_app_month_asset, intv_app_month_housing_expense, intv_exp_pay_heat_checkbox, intv_exp_pay_ac_checkbox, intv_exp_pay_electricity_checkbox, intv_exp_pay_phone_checkbox, intv_exp_pay_none_checkbox
Dim id_verif_on_file, snap_active_in_other_state, last_snap_was_exp, how_are_we_completing_the_interview
Dim cash_other_req_detail, snap_other_req_detail, emer_other_req_detail, family_cash_program, famliy_cash_notes

Dim question_1_yn, question_1_notes, question_1_verif_yn, question_1_verif_details, question_1_interview_notes
Dim question_2_yn, question_2_notes, question_2_verif_yn, question_2_verif_details, question_2_interview_notes
Dim question_3_yn, question_3_notes, question_3_verif_yn, question_3_verif_details, question_3_interview_notes
Dim question_4_yn, question_4_notes, question_4_verif_yn, question_4_verif_details, question_4_interview_notes
Dim question_5_yn, question_5_notes, question_5_verif_yn, question_5_verif_details, question_5_interview_notes
Dim question_6_yn, question_6_notes, question_6_verif_yn, question_6_verif_details, question_6_interview_notes
Dim question_7_yn, question_7_notes, question_7_verif_yn, question_7_verif_details, question_7_interview_notes
Dim question_8_yn, question_8a_yn, question_8_notes, question_8_verif_yn, question_8_verif_details, question_8_interview_notes
Dim question_9_yn, question_9_notes, question_9_verif_yn, question_9_verif_details, question_9_interview_notes
Dim question_10_yn, question_10_notes, question_10_verif_yn, question_10_verif_details, question_10_monthly_earnings, question_10_interview_notes
Dim question_11_yn, question_11_notes, question_11_verif_yn, question_11_verif_details, question_11_interview_notes
Dim pwe_selection
Dim question_12_yn, question_12_notes, question_12_verif_yn, question_12_verif_details, question_12_interview_notes
Dim question_12_rsdi_yn, question_12_rsdi_amt, question_12_ssi_yn, question_12_ssi_amt,  question_12_va_yn, question_12_va_amt, question_12_ui_yn, question_12_ui_amt, question_12_wc_yn, question_12_wc_amt, question_12_ret_yn, question_12_ret_amt, question_12_trib_yn, question_12_trib_amt, question_12_cs_yn, question_12_cs_amt, question_12_other_yn, question_12_other_amt
Dim question_13_yn, question_13_notes, question_13_verif_yn, question_13_verif_details, question_13_interview_notes
Dim question_14_yn, question_14_notes, question_14_verif_yn, question_14_verif_details, question_14_interview_notes
Dim question_14_rent_yn, question_14_subsidy_yn, question_14_mortgage_yn, question_14_association_yn, question_14_insurance_yn, question_14_room_yn, question_14_taxes_yn
Dim question_15_yn, question_15_notes, question_15_verif_yn, question_15_verif_details, question_15_interview_notes, question_15_phone_details
Dim question_15_heat_ac_yn, question_15_electricity_yn, question_15_cooking_fuel_yn, question_15_water_and_sewer_yn, question_15_garbage_yn, question_15_phone_yn, question_15_liheap_yn
Dim question_16_yn, question_16_notes, question_16_verif_yn, question_16_verif_details, question_16_interview_notes
Dim question_17_yn, question_17_notes, question_17_verif_yn, question_17_verif_details, question_17_interview_notes
Dim question_18_yn, question_18_notes, question_18_verif_yn, question_18_verif_details, question_18_interview_notes
Dim question_19_yn, question_19_notes, question_19_verif_yn, question_19_verif_details, question_19_interview_notes
Dim question_20_yn, question_20_notes, question_20_verif_yn, question_20_verif_details, question_20_interview_notes
Dim question_20_cash_yn, question_20_acct_yn, question_20_secu_yn, question_20_cars_yn
Dim question_21_yn, question_21_notes, question_21_verif_yn, question_21_verif_details, question_21_interview_notes
Dim question_22_yn, question_22_notes, question_22_verif_yn, question_22_verif_details, question_22_interview_notes
Dim question_23_yn, question_23_notes, question_23_verif_yn, question_23_verif_details, question_23_interview_notes
Dim question_24_yn, question_24_notes, question_24_verif_yn, question_24_verif_details, question_24_interview_notes
Dim question_24_rep_payee_yn, question_24_guardian_fees_yn, question_24_special_diet_yn, question_24_high_housing_yn
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_there, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five
Dim arep_name, arep_relationship, arep_phone_number, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip
Dim MAXIS_arep_name, MAXIS_arep_relationship, MAXIS_arep_phone_number, MAXIS_arep_addr_street, MAXIS_arep_addr_city, MAXIS_arep_addr_state, MAXIS_arep_addr_zip
Dim CAF_arep_name, CAF_arep_relationship, CAF_arep_phone_number, CAF_arep_addr_street, CAF_arep_addr_city, CAF_arep_addr_state, CAF_arep_addr_zip
Dim arep_complete_forms_checkbox, arep_get_notices_checkbox, arep_use_SNAP_checkbox
Dim CAF_arep_complete_forms_checkbox, CAF_arep_get_notices_checkbox, CAF_arep_use_SNAP_checkbox
Dim arep_on_CAF_checkbox, arep_action, CAF_arep_action, arep_and_CAF_arep_match, arep_authorization, arep_exists, arep_authorized
Dim signature_detail, signature_person, signature_date, second_signature_detail, second_signature_person, second_signature_date
Dim client_signed_verbally_yn, interview_date, add_to_time, update_arep, verifs_needed, verifs_selected, verif_req_form_sent_date, number_verifs_checkbox, verifs_postponed_checkbox
Dim verif_snap_checkbox, verif_cash_checkbox, verif_mfip_checkbox, verif_dwp_checkbox, verif_msa_checkbox, verif_ga_checkbox, verif_grh_checkbox, verif_emer_checkbox, verif_hc_checkbox
Dim exp_snap_approval_date, exp_snap_delays, snap_denial_date, snap_denial_explain, pend_snap_on_case, do_we_have_applicant_id
Dim family_cash_case_yn, absent_parent_yn, relative_caregiver_yn, minor_caregiver_yn
Dim disc_phone_confirmation, disc_yes_phone_no_expense_confirmation, disc_no_phone_yes_expense_confirmation, disc_homeless_confirmation, disc_out_of_county_confirmation, CAF1_rent_indicated, Verbal_rent_indicated
Dim Q14_rent_indicated, question_14_summary, disc_rent_amounts_confirmation, disc_utility_caf_1_summary, disc_utility_q_15_summary, disc_utility_amounts_confirmation

Dim confirm_resp_read, confirm_rights_read, confirm_ebt_read, confirm_ebt_how_to_read, confirm_npp_info_read, confirm_npp_rights_read
Dim confirm_appeal_rights_read, confirm_civil_rights_read, confirm_cover_letter_read, confirm_program_information_read, confirm_DV_read
Dim confirm_disa_read, confirm_mfip_forms_read, confirm_mfip_cs_read, confirm_minor_mfip_read, confirm_snap_forms_read, confirm_recap_read
Dim confirm_ievs_info_read, case_card_info, clt_knows_how_to_use_ebt_card, snap_reporting_type, next_revw_month

Dim show_pg_one_memb01_and_exp, show_pg_one_address, show_pg_memb_list, show_q_1_6
Dim show_q_7_11, show_q_14_15, show_q_21_24, show_qual, show_pg_last, discrepancy_questions, show_arep_page, expedited_determination
Dim CASH_on_CAF_checkbox, SNAP_on_CAF_checkbox, EMER_on_CAF_checkbox
Dim type_of_cash, the_process_for_cash, next_cash_revw_mo, next_cash_revw_yr
Dim the_process_for_snap, next_snap_revw_mo, next_snap_revw_yr
Dim type_of_emer, the_process_for_emer, q_12_totally_blank, q_14_totally_blank, q_15_totally_blank, q_20_totally_blank, q_24_totally_blank


'EXPEDITED DETERMINATION VARIABLES'
Dim expedited_determination_completed, determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources
Dim jobs_income_yn, busi_income_yn, unea_income_yn, cash_amount_yn, bank_account_yn, all_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, expedited_screening
Dim calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, caf_1_resources, caf_1_expenses
' Dim calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, CAF_datestamp, interview_date
Dim applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four
' Dim applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_date, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four
Dim postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn
Dim other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done
Dim rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, cash_amount
Dim previous_CAF_datestamp, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn
Dim delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn
Dim income_review_completed, assets_review_completed, shel_review_completed, note_calculation_detail


show_pg_one_memb01_and_exp	= 1
show_pg_one_address			= 2
show_pg_memb_list			= 3
show_q_1_6					= 4
show_q_7_11					= 5
show_q_12_13				= 6
show_q_14_15				= 7
show_q_16_20				= 8
show_q_21_24				= 9
show_qual					= 10
show_pg_last				= 11
discrepancy_questions		= 12
show_arep_page				= 13
expedited_determination		= 14

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
first_time_in_exp_det = True

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
verif_view = "See All Verifs"

'THE SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabbing the case number
EMConnect ""
Call check_for_MAXIS(true)
Call MAXIS_case_number_finder(MAXIS_case_number)
' CAF_datestamp = date & ""
interview_date = date & ""
show_err_msg_during_movement = ""
script_run_lowdown = ""
developer_mode = False

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
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

'Showing the case number dialog
Do
	DO
		err_msg = ""

		' EditBox 245, 50, 50, 15, CAF_datestamp
		' CheckBox 230, 80, 30, 10, "CASH", CASH_on_CAF_checkbox
		' CheckBox 270, 80, 35, 10, "SNAP", SNAP_on_CAF_checkbox
		' CheckBox 310, 80, 35, 10, "EMER", EMER_on_CAF_checkbox
		' Text 155, 55, 90, 10, "Date Application Received:"
		' GroupBox 225, 70, 125, 25, "Programs marked on CAF"

		' PushButton 205, 35, 155, 10, "NOTES - Interview Script Instructions", msg_show_instructions_btn
		' PushButton 205, 35, 155, 10, "Interview Quick Start Guide", msg_show_quick_start_guide_btn
		' PushButton 205, 35, 155, 10, "Interview FAQ", msg_show_faq_btn
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 371, 320, "Interview Script Case number dialog"
		  EditBox 75, 45, 60, 15, MAXIS_case_number
		  DropListBox 75, 65, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"MNbenefits"+chr(9)+"Combined AR for Certain Pops (DHS-3727)", CAF_form
		  EditBox 75, 85, 145, 15, worker_signature
		  DropListBox 20, 275, 335, 45, "Alert at the time you attempt to save each page of the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
		  ButtonGroup ButtonPressed
		    OkButton 260, 300, 50, 15
		    CancelButton 315, 300, 50, 15
		    PushButton 205, 20, 155, 15, "Press HERE to see what this script will do", msg_what_script_does_btn
		    PushButton 205, 35, 155, 15, "Press HERE for details on using this script", msg_script_interaction_btn
            PushButton 220, 65, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
		    PushButton 165, 175, 195, 15, "Press HERE to learn more about 'SAVE YOUR WORK'", msg_save_your_work_btn
		    PushButton 80, 245, 210, 15, "Press HERE for more details on script messaging", msg_script_messaging_btn
		    PushButton 10, 300, 50, 15, "Instructions", msg_show_instructions_btn
		    PushButton 60, 300, 70, 15, "Quick Start Guide", msg_show_quick_start_guide_btn
		    PushButton 130, 300, 30, 15, "FAQ", msg_show_faq_btn
		  Text 10, 10, 360, 10, "Start this script at the beginning of the interview and keep it running during the entire course of the interview."
		  Text 20, 50, 50, 10, "Case number:"
		  Text 10, 70, 60, 10, "Actual CAF Form:"
		  Text 10, 90, 60, 10, "Worker Signature:"
		  Text 145, 105, 105, 10, "*!*!*!*  DID YOU KNOW *!*!*!*"
		  Text 110, 120, 185, 10, "This script SAVES the information you enter as it runs!"
		  Text 75, 135, 255, 10, "This means that IF the script errors, fails, is canceled, the network goes down."
		  Text 135, 145, 125, 10, "YOU CAN GET YOUR WORK BACK!!!"
		  Text 15, 155, 345, 20, "This happens in the background, without you knowing it. In order to get your work back run the script again on the SAME DAY for the SAME CASE and it will ask if you want to restore the information - just press YES!"
		  GroupBox 10, 190, 355, 105, "How to interact with this Script"
		  Text 80, 205, 220, 10, "You should have this script running DURING the entire interview."
		  Text 90, 220, 195, 20, "You  are capturing BOTH the information writen on the form AND the verbal responses in the script fields."
		  Text 20, 265, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed > 100 Then
			err_msg = "LOOP"

			If ButtonPressed = msg_what_script_does_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20OVERVIEW.docx"
			If ButtonPressed = msg_script_interaction_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20HOW%20TO%20USE.docx"
			If ButtonPressed = interpreter_servicves_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://itwebpw026/content/forms/af/_internal/hhs/human_services/initial_contact_access/AF10196.html"
            If ButtonPressed = msg_save_your_work_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20SAVE%20YOUR%20WORK.docx"
			If ButtonPressed = msg_script_messaging_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20SCRIPT%20MESSAGING.docx"

			If ButtonPressed = msg_show_instructions_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW.docx"
			If ButtonPressed = msg_show_quick_start_guide_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20QUICK%20START%20GUIDE.docx"
			If ButtonPressed = msg_show_faq_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20FAQ.docx"
		Else
			Call validate_MAXIS_case_number(err_msg, "*")
			If no_case_number_checkbox = checked Then err_msg = ""
			' Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
			If CAF_form = "Select One:" Then err_msg = err_msg & vbCr & "* Select which form that was received that we are using for the interview."
			' If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
			IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		End If
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

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
  Text 15, 75, 210, 10, "Confrim with the resident which programs should be assessed:"
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
	'Needs to determine MyDocs directory before proceeding.
	intvw_msg_file = user_myDocs_folder & "interview message.txt"
	If user_ID_for_validation = "ERHO003" Then intvw_msg_file = user_c_drive_docs_folder & "interview message.txt"

	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(intvw_msg_file) = False then
			Set objTextStream = .OpenTextFile(intvw_msg_file, 2, true)

			'Write the contents of the text file
			objTextStream.WriteLine "While the script gathers details about the case, tell the Resident:"
			objTextStream.WriteLine ""
			objTextStream.WriteLine "- We are going to complete your required interview now."
			objTextStream.WriteLine "- I will ask you all of the questions you completed on the application:"
			objTextStream.WriteLine "  - I know this may seem repetitive but we are required to confirm the information you entered."
			objTextStream.WriteLine "  - Please answer these questions to the best of your ability."
			objTextStream.WriteLine ""
			objTextStream.WriteLine "If we cannot get all of the questions answered we cannot complete the interview."
			objTextStream.WriteLine "Unless we complete the interview, your application/recertification can not be processed."

			objTextStream.Close
		End If
	End With
	Set oExec = WshShell.Exec("notepad " & intvw_msg_file)

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

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 361, 130, "Interview Start Message"
	  ButtonGroup ButtonPressed
	    OkButton 305, 110, 50, 15
	  Text 10, 10, 220, 10, "While the script gathers details about the case, tell the Resident:"
	  Text 15, 25, 215, 10, "- We are going to complete your required interview now."
	  Text 15, 40, 250, 10, "- I will ask you all of the questions you completed on the application:"
	  Text 20, 50, 330, 10, "  - I know this may seem repetitive but we are required to confirm the information you entered."
	  Text 20, 60, 210, 10, "  - Please answer these questions to the best of your ability."
	  Text 10, 80, 275, 10, "If we cannot get all of the questions answered we cannot complete the interview."
	  Text 10, 95, 290, 10, "Unless we complete the interview, your application/recertification can not be processed."
	  Text 80, 115, 220, 10, "Press 'OK' when you have explained the interview to the resident."
	EndDialog

	dialog Dialog1

	oExec.Terminate()
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
add_verif_1_btn				= 1060
add_verif_2_btn				= 1061
add_verif_3_btn				= 1062
add_verif_4_btn				= 1063
add_verif_5_btn				= 1064
add_verif_6_btn				= 1065
add_verif_7_btn				= 1066
add_verif_8_btn				= 1070
add_verif_9_btn				= 1071
add_verif_10_btn			= 1072
add_verif_11_btn			= 1073
add_verif_12_btn			= 1074
add_verif_12_btn			= 1075
add_verif_13_btn			= 1076
add_job_btn					= 1077
add_verif_14_btn			= 1080
add_verif_15_btn			= 1081
add_verif_16_btn			= 1082
add_verif_17_btn			= 1083
add_verif_18_btn			= 1084
add_verif_19_btn			= 1085
add_verif_20_btn			= 1090
add_verif_21_btn			= 1091
add_verif_22_btn			= 1092
add_verif_23_btn			= 1093
add_verif_24_btn			= 1094
add_verif_jobs_btn			= 1095
clear_job_btn				= 1100
' open_r_and_r_button 		= 1200
caf_page_one_btn			= 1300
caf_addr_btn				= 1400
caf_membs_btn				= 1500
caf_q_1_6_btn				= 1600
caf_q_7_11_btn				= 1700
caf_q_12_13_btn				= 1800
caf_q_14_15_btn				= 1900
caf_q_16_20_btn				= 2000
caf_q_21_24_btn				= 2100
caf_qual_q_btn				= 2200
caf_last_page_btn			= 2300
finish_interview_btn		= 2400
exp_income_guidance_btn 	= 2500
discrepancy_questions_btn	= 2600
open_hsr_manual_transfer_page_btn = 2610
incomplete_interview_btn	= 2700
verif_button				= 2800
q_12_all_no_btn				= 2900
q_14_all_no_btn				= 3000
expedited_determination_btn	= 3010
return_btn 					= 900

btn_placeholder = 4000
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	JOBS_ARRAY(jobs_edit_btn, each_job) = btn_placeholder
	btn_placeholder = btn_placeholder + 1
next
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
				' MsgBox page_display
				' MsgBox update_arep & " - before define dlg"
				Dialog1 = Empty
				call define_main_dialog

				err_msg = ""

				prev_page = page_display
				previous_button_pressed = ButtonPressed
				' MsgBox update_arep & " - before display dlg"

				dialog Dialog1
				save_your_work
				cancel_confirmation
				' MsgBox  HH_MEMB_ARRAY(0).ans_imig_status
				Call review_information
				Call assess_caf_1_expedited_questions(expedited_screening)
				Call review_for_discrepancies
				Call verification_dialog
				Call check_for_errors(interview_questions_clear)
				' If show_err_msg_during_movement = FALSE AND ButtonPressed <> finish_interview_btn Then err_msg = ""
                If ButtonPressed = interpreter_servicves_btn Then
                    run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://itwebpw026/content/forms/af/_internal/hhs/human_services/initial_contact_access/AF10196.html"
                Else
                    Call display_errors(err_msg, False, show_err_msg_during_movement)
                End If
				' If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

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
'TODO - add a check_for_MAXIS here once GH 1166 is done and the dialog call doesn't break the interview

If relative_caregiver_yn = "Yes" Then absent_parent_yn = "Yes"
exp_pregnant_who = trim(exp_pregnant_who)
If exp_pregnant_who = "Select or Type" Then exp_pregnant_who = ""

for each_member = 0 to UBound(HH_MEMB_ARRAY, 2)
	If HH_MEMB_ARRAY(id_verif, each_member) = "Found in SOLQ/SMI" Then HH_MEMB_ARRAY(id_verif, each_member) = "Identity verified per Verify MN interface"
next

interview_complete = True
If ButtonPressed = incomplete_interview_btn Then
	' MsgBox "ARE YOU SURE?"
	interview_complete = False
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 436, 200, "Interview Incomplete"
			  EditBox 10, 70, 420, 15, interview_incomplete_reason
			  CheckBox 15, 110, 250, 10, "Check Here to Create a CASE:NOTE with the detail that was gathered.", create_incomplete_note_checkbox
			  ' CheckBox 15, 125, 265, 10, "Check here to create a document with the partial notes that exist at this point.", create_incomplete_doc_checkbox
			  EditBox 10, 160, 420, 15, incomplete_interview_notes
			  ButtonGroup ButtonPressed
			    OkButton 380, 180, 50, 15
			  GroupBox 5, 10, 425, 40, "Incompleting an Interview"
			  Text 15, 20, 405, 25, "We make every attempt to complete the entire interview requirement when we are in contact with the resident. Sometimes this becomes impossible and if we are unable to gather all required information, we must INCOMPLETE the interview. Every attempt should be made to complete the interview first."
			  Text 10, 60, 120, 10, "Reason the Interview is Incomplete"
			  GroupBox 5, 95, 425, 45, "Options "
			  Text 10, 150, 75, 10, "Additional Notes"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			interview_incomplete_reason = trim(interview_incomplete_reason)
			incomplete_interview_notes = trim(incomplete_interview_notes)

			If interview_incomplete_reason = "" Then err_msg = err_msg & vbCr & "* Explain why the interview is incomplete."

			If err_msg <> "" Then MsgBox "*****     NOTICE     *****" & vbCr & "Please resolve to continue:" & vbCr & err_msg

		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False

	If create_incomplete_doc_checkbox = checked Then

	End If

	If create_incomplete_note_checkbox = checked Then
		Call write_verification_CASE_NOTE(create_verif_note)
		Call write_interview_CASE_NOTE
		PF3
	End If

	Call start_a_blank_case_note

	Call write_variable_in_CASE_NOTE("INTERVIEW INCOMPLETE - Attempt made but additional details needed")

	Call write_variable_in_CASE_NOTE("Interview attempted on: " & interview_date)
	If create_incomplete_doc_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Document added to Case File with information that was gathered during this partial interview.")
	If create_incomplete_note_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Previous CASE:NOTE has details of information what was gathered during this partial interview.")
	Call write_bullet_and_variable_in_CASE_NOTE("Reason Interview Incomplete", interview_incomplete_reason)
	Call write_bullet_and_variable_in_CASE_NOTE("Additional Notes", incomplete_interview_notes)

	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	time_spent = ((timer - start_time) + add_to_time)/60
	time_spent = Round(time_spent, 2)
	end_msg ="INCOMPLETE INTERVIEW run finished." & vbCr & vbCr & "You spent " & time_spent & " minutes on this interview."
	If create_incomplete_doc_checkbox = checked Then end_msg = end_msg & vbCr & " - Doc created to add to ECF."
	If create_incomplete_note_checkbox = checked Then end_msg = end_msg & vbCr & " - NOTE with gathered information created."

	Call script_end_procedure(end_msg)
End If
'Navigate back to self and to EDRS
Back_to_self
CALL navigate_to_MAXIS_screen("INFC", "EDRS")
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

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
    					' GroupBox 15, y_pos - 15, 520, 50, "MEMB " & HH_MEMB_ARRAY(ref_number, the_memb) & " - " & HH_MEMB_ARRAY(full_name_const, the_memb)
    					Text 30, y_pos + 20, 45, 10, "EDRs Notes:"
    		  		    EditBox 80, y_pos + 15, 450, 15, HH_MEMB_ARRAY(edrs_notes, the_memb)
    					y_pos = y_pos + 20
    				End If
    				' If HH_MEMB_ARRAY(edrs_match, the_memb) = FALSE Then GroupBox 15, y_pos - 15, 520, 30, "MEMB XX - MEMBER NAME"
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

If MFIP_orientation_assessed_and_completed = False Then
    If cash_request = True and the_process_for_cash = "Application" and type_of_cash = "Family" Then
        Call complete_MFIP_orientation(HH_MEMB_ARRAY, ref_number, full_name_const, age, memb_is_caregiver, cash_request_const, hours_per_week_const, exempt_from_ed_const, comply_with_ed_const, orientation_needed_const, orientation_done_const, orientation_exempt_const, exemption_reason_const, emps_exemption_code_const, choice_form_done_const, orientation_notes, family_cash_program)
        For caregiver = 0 to UBound(HH_MEMB_ARRAY, 2)
            If HH_MEMB_ARRAY(memb_is_caregiver, caregiver) = True and HH_MEMB_ARRAY(orientation_needed_const, caregiver) = True and HH_MEMB_ARRAY(orientation_done_const, caregiver) = False and HH_MEMB_ARRAY(orientation_exempt_const, caregiver) = False Then
                Call start_a_blank_CASE_NOTE

                Call write_variable_in_CASE_NOTE("MF Orientation NOT COMPLETED for " & HH_MEMB_ARRAY(full_name_const, caregiver))
                Call write_variable_in_CASE_NOTE("Interview completed but could not complete the MFIP Orientation at the time of the interview.")
                Call write_variable_in_CASE_NOTE("* MFIP Orientation is still needed for " & HH_MEMB_ARRAY(full_name_const, caregiver))
                Call write_variable_in_CASE_NOTE(HH_MEMB_ARRAY(full_name_const, caregiver) & " did not meet an exemption from completing an MFIP Orientation")
                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE(worker_signature)
                PF3
                call back_to_SELF

            End If
        Next
    End If
End If
MFIP_orientation_assessed_and_completed = True
save_your_work

'CLIENT RESPONSIBILITEIS
If left(confirm_resp_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Responsibilities Discussed"+chr(9)+"No, I could not complete this", confirm_resp_read
			  GroupBox 10, 25, 530, 335, "Rights and Responsibilities Text"
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
				PushButton 430, 22, 100, 13, "Open DHS 4163", open_r_and_r_btn
			  Text 10, 10, 160, 10, "REVIEW the information listed here to the resident:"
			  Text 20, 35, 505, 35, "Note: Cash on an Electronic Benefit Transfer (EBT) card is provided to help families meet their basic needs, including: food, shelter, clothing, utilities and transportation. These funds are provided until families can support themselves. It is illegal for an EBT user to buy or attempt to buy tobacco products or alcohol with the EBT card. If you do, it is fraud and you will be removed from the program. Do not use an EBT card at a gambling establishment or retail establishment, which provides adult-orientated entertainment in which performers disrobe or perform in an unclothed state for entertainment."
			  Text 20, 70, 275, 50, "- If you receive cash assistance and/or child care assistance, you must report changes which may affect your benefits to the county agency within 10 days after the change has occurred. If you receive Supplemental Nutrition Assistance Program (SNAP) benefits, report changes by the 10th of the month following the month of the change. Each program may have different requirements for reporting changes. Talk to your caseworker about what you must report."

			  Text 20, 120, 275, 10, "You may be required to report changes in:"
			  Text 20, 130, 275, 20, "-Employment - starting or stopping a job or business; change in hours, earnings or expenses"
			  Text 20, 150, 275, 25, "- Income - receipt or change in child support, Social Security, veteran benefits, unemployment insurance, inheritance or insurance benefits"
			  Text 20, 170, 275, 20, "- Property - purchase, sale or transfer of a house, car or other items of value, or if you receive an inheritance or settlement"
			  Text 20, 190, 275, 20, "- Household - When a person dies or becomes disabled, moves in or out of your home or temporarily leaves; pregnancy; birth of a child."
			  Text 20, 210, 275, 10, "- Citizenship or immigration status"
			  Text 20, 220, 275, 10, "- Address"
			  Text 20, 230, 275, 10, "- Housing costs and/or rent subsidy"
			  Text 20, 240, 275, 10, "- Utility costs"
			  Text 20, 250, 275, 10, "- Filing a lawsuit"
			  Text 20, 260, 275, 10, "- Absent parent custody or visits"
			  Text 20, 270, 275, 10, "- Drug felony conviction"
			  Text 20, 280, 275, 10, "- Marriage, separation or divorce"
			  Text 20, 290, 275, 10, "- School attendance"
			  Text 20, 300, 275, 10, "- Health insurance coverage and premiums"
			  Text 20, 315, 275, 20, "Note: If you change child care providers, you must tell your child care worker and provider at least 15 days before the change goes into effect."

			  Text 15, 335, 520, 10, "If you have any questions or are unsure about any reporting rules, contact your worker. If your worker is not available, leave a message so the worker can get back to you."

			  Text 310, 70, 225, 35, "- The county, state or federal agency may check any of the information you provide. To obtain some forms of information we must have your signed consent. If you don't allow the county to confirm your information, you might not receive assistance."
			  Text 310, 105, 225, 35, "- If you give us information you know is untrue, withhold information or do not report as required, or we discover your information is untrue, you may be investigated for fraud. This may result in you being disqualified from receiving benefits, charged criminally, or both."
			  Text 310, 140, 225, 50, "- The state or federal quality control agency may randomly choose your case for review. They will review statements you provided and will check to see if your eligibility was figured correctly. The state may seek information from other sources and will inform you about any contact they intend to make. If you do not cooperate, your benefits may stop."
			  Text 310, 195, 225, 10, "Cooperation requirements:"
			  Text 310, 205, 225, 45, "- If the county approves you for the Minnesota Family Investment Program (MFIP) or the Diversionary Work Program (DWP), you must cooperate with employment services, unless you are exempt. You must develop and sign an employment plan or your DWP application will be denied."
			  Text 310, 250, 225, 55, "- To receive MFIP, DWP, and/or child care assistance, you must cooperate with child support enforcement for all children in your household. You have the right to claim 'good cause' for not cooperating with child support enforcement. Yo must assign your child support to the state of Minnesota for all eligible children. If you do not cooperate or assign your child support, benefits will be denied or terminated."
			  Text 310, 305, 225, 30, "After the county approves your MFIP or DWP, if you receive child support directly from the noncustodial parent, you must report it to your worker."

			  Text 10, 370, 210, 10, "Confirm you have reviewed resident responsibilities:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_resp_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_r_and_r_btn Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False
End If
save_your_work

'CLIENT RIGHTS
If left(confirm_rights_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 160, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Rights Discussed"+chr(9)+"No, I could not complete this", confirm_rights_read
			  GroupBox 10, 25, 530, 335, "Rights and Responsibilities Text"
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
				PushButton 430, 22, 100, 13, "Open DHS 4163", open_r_and_r_btn
			  Text 10, 10, 160, 10, "REVIEW the information listed here to the resident:"

			  Text 275, 35, 150, 10, "Your Rights"

			  Text 20, 50, 275, 30, "- Your right to privacy. Your private information, including your health information, is protected by state and federal laws. Your worker has given you a Notice of Privacy Practices (DHS-3979) information sheet explaining these rights."
			  Text 20, 85, 275, 10, "- You have the right to reapply at any time if your benefits stop."
			  Text 20, 95, 275, 20, "- You have the right to receive a paper OR electronic copy of your SNAP application."
			  Text 20, 105, 275, 25, "- You have the right to know why, if we have not processed your application within:"
			  Text 30, 115, 265, 20, "- 30 days for cash, SNAP and child care assistance"
			  Text 30, 125, 265, 20, "- 60 days for cash related to disability."
			  Text 20, 135, 275, 25, "- You have the right to know the rules of the program you are applying for and for the agency to tell you how your benefit amount was figured."
			  Text 20, 155, 275, 10, "- You have the right to choose where and with whom you live."
			  Text 20, 165, 275, 45, "- Expenses. You have the right to report expenses such as shelter, utilities, child care, child support or medical costs. These expenses may affect the amount of Supplemental Nutrition Assistance Program (SNAP) benefits that you receive. Failure to report or verify certain expenses listed will be a statement by your household that you do not want a deduction for the unreported expenses."

			  Text 310, 50, 225, 35, "For SNAP, you may appeal within 90 days by writing or calling the county or the State Appeals Office. You may represent yourself at the hearing, or you may have someone (an attorney, relative, friend or another person) speak for you."
			  Text 310, 90, 225, 50, "If you wish your assistance to continue until the hearing, you must appeal before the date of the proposed action or within 10 days after the date the agency notice was mailed, whichever is later. Ask your county or tribal worker to explain how the timing of your appeal could affect your present or future assistance."
			  Text 310, 140, 225, 20, "- Access to free legal services. Contact your worker for information on free legal services."
			  Text 310, 165, 225, 80, "- Appeal rights. If you are unhappy with the action taken or feel the agency did not act on your request for assistance, you may appeal. For cash, child care assistance and health care, you may appeal within 30 days from the date you receive the notice by writing to the county or tribal agency, or directly to the State Appeals Office at the Minnesota Department of Human Services, PO Box 64941, St. Paul, MN 55164-0941. (If you show good cause for not appealing your cash and health care within 30 days, the agency can accept your appeal for up to 90 days from the date you receive the notice.)"

			  Text 10, 370, 150, 10, "Confirm you have reviewed resident rights:"
			EndDialog

			dialog Dialog1
	 		cancel_confirmation

			If confirm_rights_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_r_and_r_btn Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

case_number_last_digit = right(MAXIS_case_number, 1)
case_number_last_digit = case_number_last_digit * 1
If case_number_last_digit = 4 Then snap_day_of_issuance = "4th"
If case_number_last_digit = 5 Then snap_day_of_issuance = "5th"
If case_number_last_digit = 6 Then snap_day_of_issuance = "6th"
If case_number_last_digit = 7 Then snap_day_of_issuance = "7th"
If case_number_last_digit = 8 Then snap_day_of_issuance = "8th"
If case_number_last_digit = 9 Then snap_day_of_issuance = "9th"
If case_number_last_digit = 0 Then snap_day_of_issuance = "10th"
If case_number_last_digit = 1 Then snap_day_of_issuance = "11th"
If case_number_last_digit = 2 Then snap_day_of_issuance = "12th"
If case_number_last_digit = 3 Then snap_day_of_issuance = "13th"
If case_number_last_digit MOD 2 = 1 Then cash_day_of_issuance = "2nd to last day"		'ODD Number
If case_number_last_digit MOD 2 = 0 Then cash_day_of_issuance = "last day"		'EVEN Number
If cash_type = "ADULT" Then cash_day_of_issuance = "first day"

'EBT RESPONSIBILITIES AND USAGE
If left(confirm_ebt_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  ComboBox 310, 45, 225, 45, "Select or Type"+chr(9)+"Yes - I have my card."+chr(9)+"No - I used to but I've lost it."+chr(9)+"No - I never had a card for this case"+chr(9)+case_card_info, case_card_info
			  DropListBox 310, 75, 225, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", clt_knows_how_to_use_ebt_card
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! EBT Basics Discussed"+chr(9)+"No, I could not complete this", confirm_ebt_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 10, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 25, 530, 335, "EBT Information"
			  Text 20, 35, 275, 10, "For Cash and Supplemental Nutrition Assistance Program (SNAP) benefits:"
			  Text 30, 45, 265, 25, "- Each time you use your Electronic Benefits Transfer (EBT) card or sign your check, you state that you have informed the county or tribal agency about any changes in your situation that may affect your benefits."
			  Text 30, 75, 265, 25, "- Each time your EBT card is used, we assume you have received your cash or SNAP benefits, unless you reported your card lost or stolen to the county or tribal agency."

			  Text 20, 105, 275, 25, "The standard way to get your benefits to you is through issuance on an EBT card. For cash benefits, there may be other options such as a vendor payment or direct deposit. If you want more information about these options, please let us know."

			  Text 20, 140, 275, 10, "EBT card balances and information can be found:"
			  Text 30, 150, 265, 10, "- Call customer service, 24 hours a day / 7 days a week - Toll-free: 888-997-2227"
			  Text 30, 160, 265, 25, "- Go to www.ebtEDGE.com - Under EBT Cardholders, click on 'More Information' and log in using your user ID and password."
			  ' Text 20, 105, 275, 25, ""

			  GroupBox 10, 190, 290, 75, "Your EBT Issuances"
			  Text 20, 205, 275, 10, "If approved, your SNAP benefits will regularly be issued on the " & snap_day_of_issuance & " of the month."
			  Text 20, 220, 275, 10, "If approved, your CASH benefits will regularly be issued on the " & cash_day_of_issuance & " of the month."
			  Text 20, 235, 275, 20, "*** Due to processing changes or delay in receipt of information issuances days may change, you should access EBT information directly to ensure benefits are available."


			  Text 310, 35, 225, 10, "Do you already have an EBT card for this case?"

			  Text 310, 65, 225, 10, "Do you know how to use an EBT card?"

			  Text 10, 370, 210, 10, "Confirm you have reviewed EBT Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_ebt_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."
			If confirm_ebt_read = "YES! EBT Basics Discussed" Then
				If case_card_info = "Select or Type" or trim(case_card_info) = "" Then err_msg = err_msg & vbNewLine & "* Since you have discussed EBT Information, indicate if the resident has an EBT Card for this case."
				If clt_knows_how_to_use_ebt_card = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since you have discussed EBT Information, indicate if the resident knows how to use their EBT Card."
			End If
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

If clt_knows_how_to_use_ebt_card = "No" then
	If left(confirm_ebt_how_to_read, 4) <> "YES!" Then
		Do
			Do
				err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
				  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! EBT Detail Discussed"+chr(9)+"No, I could not complete this", confirm_ebt_how_to_read
				  ButtonGroup ButtonPressed
				    PushButton 465, 365, 80, 15, "Continue", continue_btn
				  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
				  GroupBox 10, 15, 530, 340, "How to Use Your Minnesota EBT Card"
				  Text 185, 25, 345, 10, "Your EBT card is a safe, convenient and easy way for you to get your cash and food benefits each month."
				  Text 10, 370, 210, 10, "Confirm you have reviewed How to Use EBT Information:"
				  ButtonGroup ButtonPressed
				    PushButton 440, 5, 100, 13, "Open DHS 3315A", open_ebt_brochure_btn
				  Text 20, 30, 65, 10, "How to get a card:"
				  Text 25, 40, 305, 10, "- Your first card will be mailed to you within 2 business days of your benefits being approved."
				  Text 25, 50, 130, 10, "- Replacement cards are also mailed."
				  Text 40, 60, 170, 10, "Call 1-888-997-2227 to request a replacement card"
				  Text 40, 70, 170, 10, "Cards take about 5 business days to arrive."
				  Text 40, 80, 275, 10, "There is a $2 charge for all replacement cards, which is reduced from your benefit."
				  Text 25, 90, 230, 20, "NOTE: If you have cash benefits, you will be issued a card that has your name on it. SNAP only cases to not have names on the EBT card."
				  Text 20, 115, 85, 10, "Where to use your card:"
				  Text 25, 125, 120, 10, "At a store 'point-of-sale' machine."
				  Text 25, 135, 75, 10, "At an ATM (Cash Only)"
				  Text 25, 145, 140, 10, "At a check cashing business (Cash Only)"
				  Text 365, 45, 80, 10, "Keep your card safe"
				  Text 375, 55, 120, 10, "Lost benefits will not be replaced."
				  Text 375, 65, 155, 15, "Do not leave your card lying around or lose it, treat it like a debit card or cash."
				  Text 365, 90, 110, 10, "Do not throw your card away"
				  Text 375, 105, 150, 20, "The same card will be used every month for as long as you have benefits."
				  Text 375, 130, 155, 20, "Even if your cases closes and reopens in the future the same card may be used."
				  Text 365, 155, 145, 10, "Misuse of your EBT Card is Unlawful"
				  Text 370, 170, 160, 20, "- Selling your card or PIN to others may result in criminal charges and your benefits may end."
				  Text 370, 190, 165, 20, "- Attempting to buy tobacco products or alcoholic beverages with your EBT Card is considered fraud."
				  Text 370, 210, 165, 20, "- Repeated loss of your card may cause a fraud investigation to be opened on you."
				  Text 20, 165, 105, 10, "How to get or change your PIN:"
				  Text 25, 180, 135, 10, "- Call customer service at 888-997-2227"
				  Text 25, 190, 165, 10, "- Visit your county or tribal human services office"
				  Text 25, 200, 195, 10, "- Visit the ebtEDGE cardholder portal www.ebtEDGE.com"
				  Text 25, 210, 195, 20, "- Access the ebtEDGE mobile application, www.FISGLOBal.COM/EBTEDGEMOBILE"
				  Text 20, 230, 145, 20, "4 failed attepts to enter your PIN will lock your card until 12:01 am the next day."
				  Text 20, 255, 185, 10, "Register to receive EBT Information by Text Message"
				  Text 35, 325, 135, 10, "- Current Balance (text 'BAL' to 42265)"
				  Text 35, 335, 145, 10, "- Last 5 transactions  (text 'MINI' to 42265)"
				  Text 25, 265, 135, 10, "1. Go to www.ebtEDGE.com and log in"
				  Text 25, 275, 80, 10, "2. Select 'EBT Account'"
				  Text 25, 285, 205, 10, "3. Select 'Messaging Registration' under the Account Services menu"
				  Text 25, 295, 140, 10, "4. Enter your mobile (cell) phone number."
				  Text 25, 305, 230, 10, "5. Check the box next to SMS Balance, then click the 'Update' button."
				  Text 25, 315, 190, 10, "6. Use the same mobil number and text for information:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If confirm_ebt_how_to_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

				If ButtonPressed = open_ebt_brochure_btn Then
					err_msg = "LOOP"
					run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3315A-ENG"
				End If

				IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE
	End If
	save_your_work
End If


'NOTICE OF PRIVACY PRACTICES
If left(confirm_npp_info_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Notice of Privacy Information Discussed"+chr(9)+"No, I could not complete this", confirm_npp_info_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
				PushButton 440, 5, 100, 13, "Open DHS 3979", open_npp_doc
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Notice of Privacy Practices - About the Information you give us"
			  Text 20, 25, 505, 35, "This notice tells how private information about you may be used and disclosed and how you can get this information. Please review it carefully."
			  Text 15, 35, 275, 10, "Why do we ask for this information?"
			  Text 15, 45, 275, 10, "In order to determine whether and how we can help you, we collect information:"
			  Text 17, 55, 3, 10, "-"
			  Text 20, 55, 275, 10, "To tell you apart from other people with the same or similar name"
			  Text 17, 65, 3, 10, "-"
			  Text 20, 65, 275, 10, "To decide what you are eligible for"
			  Text 17, 75, 3, 10, "-"
			  Text 20, 75, 275, 20, "To help you get medical, mental health, financial or social services and decide if you can pay for some services"
			  Text 17, 95, 3, 10, "-"
			  Text 20, 95, 275, 10, "To decide if you or your family need protective services"
			  Text 17, 105, 3, 10, "-"
			  Text 20, 105, 275, 10, "To decide about out-of-home care and in-home care for you or your children"
			  Text 17, 115, 3, 10, "-"
			  Text 20, 115, 275, 10, "To investigate the accuracy of the information in your application"
			  Text 15, 125, 275, 20, "After we have begun to provide services or support to you, we may collect additional information:"
			  Text 17, 145, 3, 10, "-"
			  Text 20, 145, 275, 10, "To make reports, do research, do audits, and evaluate our programs"
			  Text 17, 155, 3, 10, "-"
			  Text 20, 155, 275, 10, "To investigate reports of people who may lie about the help they need"
			  Text 17, 165, 3, 10, "-"
			  Text 20, 165, 275, 20, "To collect money from other agencies, like insurance companies, if they should pay for your care"
			  Text 17, 180, 3, 10, "-"
			  Text 20, 180, 275, 10, "To collect money from the state or federal government for help we give you."
			  Text 17, 190, 3, 10, "-"
			  Text 20, 190, 275, 20, "When your or your family's circumstances change and you are required to report the change (see Client Responsibilities and Rights - DHS-4163)"
			  Text 15, 210, 275, 10, "Why do we ask you for your Social Security number?"
			  Text 20, 220, 275, 75, "We need your Social Security number to give you medical assistance, some kinds of financial help, or child support enforcement services (42 CFR 435.910 [2006]; Minn. Stat. 256D.03, subd.3(h); Minn. Stat.256L.04, subd. 1a; 45 CFR 205.52 [2001]; 42 USC 666; 45 CFR 303.30 [2001]). We also need your Social Security Number to verify identity and prevent duplication of state and federal benefits. Additionally, your Social Security Number is used to conduct computer data matches with collaborative, nonprofit and private agencies to verify income, resources, or other information that may affect your eligibility and/or benefits."
			  Text 20, 285, 275, 10, "You do not have to give us the Social Security Number:"
			  Text 22, 295, 3, 10, "-"
			  Text 25, 295, 275, 10, "For persons in your home who are not applying for coverage"
			  Text 22, 305, 3, 10, "-"
			  Text 25, 305, 275, 10, "If you have religious objections"
			  Text 22, 315, 3, 10, "-"
			  Text 25, 315, 500, 10, "If you are not a United States citizen and are applying for Emergency Medical Assistance only"
			  Text 22, 325, 3, 10, "-"
			  Text 25, 325, 500, 20, "If you are from another country, in the United States on a temporary basis and do not have permission from the United States Citizenship and Immigration Services to live in the United States permanently"
			  Text 22, 342, 3, 10, "-"
			  Text 25, 342, 500, 10, "If you are living in the United States without the knowledge or approval of the U.S. Citizenship and Immigration Services."
			  Text 305, 35, 225, 10, "Do you have to answer the questions we ask?"
			  Text 310, 45, 240, 45, "You do not have to give us your personal information. Without the information, we may not be able to help you. If you give us wrong information on purpose, you can be investigated and charged with fraud."
			  Text 305, 75, 225, 10, "With whom may we share information?"
			  Text 305, 85, 225, 35, "We will only share information about you as needed and as allowed or required by law. We may share your information with the following agencies or persons who need the information to do their jobs:"
			  Text 307, 110, 3, 10, "-"
			  Text 310, 110, 225, 35, "Employees or volunteers with other state, county, local, federal, collaborative, nonprofit and private agencies"
			  Text 307, 130, 3, 10, "-"
			  Text 310, 130, 225, 35, "Researchers, auditors, investigators, and others who do quality of care reviews and studies or commence prosecutions or legal actions related to managing the human services programs."
			  Text 307, 155, 3, 10, "-"
			  Text 310, 155, 225, 35, "Court officials, county attorney, attorney general, other law enforcement officials, child support officials, and child protection and fraud investigators"
			  Text 307, 180, 3, 10, "-"
			  Text 310, 180, 225, 10, "Human services offices, including child support enforcement offices"
			  Text 307, 190, 3, 10, "-"
			  Text 310, 190, 225, 20, "Governmental agencies in other states administering public benefits programs"
			  Text 307, 210, 3, 10, "-"
			  Text 310, 210, 225, 20, "Health care providers, including mental health agencies and drug and alcohol treatment facilities"
			  Text 307, 230, 3, 10, "-"
			  Text 310, 230, 225, 20, "Health care insurers, health care agencies, managed care organizations and others who pay for your care"
			  Text 307, 250, 3, 10, "-"
			  Text 310, 250, 225, 10, "Guardians, conservators or persons with power of attorney"
			  Text 307, 260, 3, 10, "-"
			  Text 310, 260, 225, 20, "Coroners and medical investigators if you die and they investigate your death"
			  Text 307, 280, 3, 10, "-"
			  Text 310, 280, 225, 20, "Credit bureaus, creditors or collection agencies if you do not pay fees you owe to us for services"
			  Text 307, 300, 3, 10, "-"
			  Text 310, 300, 225, 10, "Anyone else to whom the law says we must or can give the information"
			  Text 10, 370, 210, 10, "Confirm you have reviewed Privacy Practices Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_npp_info_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_npp_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

If left(confirm_npp_rights_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
	  		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Notice of Privacy Rights Discussed"+chr(9)+"No, I could not complete this", confirm_npp_rights_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
				PushButton 440, 5, 100, 13, "Open DHS 3979", open_npp_doc
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Notice of Privacy Practices - Rights"
			  Text 20, 25, 505, 35, "This notice tells how private information about you may be used and disclosed and how you can get this information. Please review it carefully."
			  Text 15, 40, 275, 10, "What are your rights regarding the information we have about you?"
			  Text 17, 50, 3, 10, "-"
			  Text 20, 50, 275, 20, "You and people you have given permission to may see and copy private information we have about you. You may have to pay for the copies."
			  Text 17, 70, 3, 10, "-"
			  Text 20, 70, 275, 40, "You may question if the information we have about you is correct. Send your concerns in writing. Tell us why the information is wrong or not complete. Send your own explanation of the information you do not agree with. We will attach your explanation any time information is shared with another agency."
			  Text 17, 110, 3, 10, "-"
			  Text 20, 110, 275, 35, "You have the right to ask us in writing to share information with you in a certain way or in a certain place. For example, you may ask us to send health information to your work address instead of your home address. If we find that your request is reasonable, we will grant it."
			  Text 17, 150, 3, 10, "-"
			  Text 20, 150, 275, 20, "You have the right to ask us to limit or restrict the way that we use or disclose your information, but we are not required to agree to this request."
			  Text 17, 170, 3, 10, "-"
			  Text 20, 170, 275, 20, "If you do not understand the information, ask your worker to explain it to you. You can ask the Minnesota Department of Human Services for another copy of this notice."

			  Text 15, 200, 150, 10, "What privacy rights do children have?"
			  Text 20, 215, 490, 50, "If you are under 18, when parental consent for medical treatment is not required, information will not be shown to parents unless the health care provider believes not sharing the information would risk your health. Parents may see other information about you and let others see this information, unless you have asked that this information not be shared with your parents. You must ask for this in writing and say what information you do not want to share and why. If the agency agrees that sharing the information is not in your best interest, the information will not be shared with your parents. If the agency does not agree, the information may be shared with your parents if they ask for it."
			  Text 15, 270, 275, 10, "What if you believe your privacy rights have been violated?"
			  Text 20, 285, 490, 20, "If you think that the Minnesota Department of Human Services has violated your privacy rights, you may send a written complaint to the U.S. Department of Health and Human Services to the address below:"
			  Text 20, 305, 275, 10, "Minnesota Department of Human Services"
			  Text 20, 315, 275, 10, "Attn: Privacy Official"
			  Text 20, 325, 275, 10, "PO Box 64998"
			  Text 20, 335, 275, 10, "St. Paul, MN 55164-0998"

			  Text 305, 40, 225, 10, "What are our responsibilities?"
			  Text 307, 50, 3, 10, "-"
			  Text 310, 50, 225, 20, "We must protect the privacy of your private information according to the terms of this notice."
			  Text 307, 70, 3, 10, "-"
			  Text 310, 70, 225, 40, "We may not use your information for reasons other than the reasons listed on this form or share your information with individuals and agencies other than those listed on this form unless you tell us in writing that we can."
			  Text 307, 110, 3, 10, "-"
			  Text 310, 110, 225, 40, "We must follow the terms of this notice, but we may change our privacy policy because privacy laws change. We will put changes to our privacy rules on our website at: http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
			  ' ButtonGroup ButtonPressed
			  '   PushButton 310, 150, 100, 13, "Open DHS 3979", open_npp_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Privacy Practices Rights:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_npp_rights_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_npp_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

'NOTICE ABOUT IEVS
If left(confirm_ievs_info_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! IEVS Information Discussed"+chr(9)+"No, I could not complete this", confirm_ievs_info_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "IEVS Information"
			  Text 15, 25, 275, 10, "What is the Income and Eligibility Verification System (IEVS)?"
			  Text 20, 35, 275, 20, "The government has a way to check income. It is the 'Income and Eligibility Verification System' (IEVS)."
			  Text 20, 55, 275, 30, "The law has us check your income with other agencies. We have to check income for all who ask for or get cash assistance, Supplemental Nutrition Assistance Program (SNAP) benefits or Medical Assistance (MA). This includes your children."
			  Text 20, 85, 275, 30, "We need Social Security Numbers (SSN) for anyone wanting help. If you have no SSN, you must apply for one. Apply with your county human services agency. You must report all SSNs to your worker."

			  Text 15, 115, 275, 10, "Agencies we get information from. We must trade facts with these agencies:"
			  Text 17, 125, 3, 10, "-"
			  Text 20, 125, 275, 10, "United States Social Security Administration (SSA)"
			  Text 17, 135, 3, 10, "-"
			  Text 20, 135, 275, 10, "United States Internal Revenue Service (IRS)"
			  Text 17, 145, 3, 10, "-"
			  Text 20, 145, 275, 10, "Minnesota Department of Employment and Economic Development (DEED)"
			  Text 17, 155, 3, 10, "-"
			  Text 20, 155, 275, 10, "Minnesota Office of Child Support Division"
			  Text 17, 165, 3, 10, "-"
			  Text 20, 165, 275, 10, "Agencies in other states that manage:"
			  Text 17, 175, 3, 10, "-"
			  Text 20, 175, 275, 10, "Unemployment Insurance"
			  Text 17, 185, 3, 10, "-"
			  Text 20, 185, 275, 10, "Cash assistance/SNAP/MA"
			  Text 17, 195, 3, 10, "-"
			  Text 20, 195, 275, 10, "Child support"
			  Text 17, 205, 3, 10, "-"
			  Text 20, 205, 275, 10, "SSI state supplements"
			  Text 15, 215, 275, 30, "These agencies have the right to get certain facts from us about you. They have to use those facts for programs like RSDI, child support, cash assistance, SNAP, MA, Unemployment Insurance, and SSI."

			  Text 15, 230, 275, 10, "Your duty to report"
			  Text 20, 240, 275, 10, "You must report all of your income and assets."
			  Text 20, 250, 275, 20, "You must still report all of your income, assets and other information on redetermination forms we send you.  "
			  Text 20, 270, 275, 20, "You must help the county agency check your income, assets and health insurance. IEVS is one way of proving your income, assets and health insurance amounts."
			  Text 15, 290, 275, 10, "What if you do not help"
			  Text 20, 300, 275, 20, "You must help us check your income, assets and health insurance to get cash assistance, SNAP and MA. If you don't, you and your family will not get help."

			  Text 120, 330, 380, 20, "Legal Authority - IEVS - 7 CFR, parts 271, 272, 273, 275; 42 CFR, parts 431, 435; 45 CFR, parts 205, 206, 233 - Work Reporting - Minnesota Statutes Section 256.998, Subd. 10"

			  Text 305, 25, 225, 10, "What facts will we get? How will we use them?"
			  Text 305, 35, 225, 40, "We check with other agencies about your income, assets and health insurance. If you didn't tell us about all of your income or assets, we will refigure your aid. Your aid might go lower or stop. If you get aid you should not be getting, we may use these facts in civil or criminal lawsuits."
			  Text 305, 75, 225, 40, "We will tell you if facts from other agencies are not the same as the facts you gave us. We will tell you what facts we got, the kind of income or assets, and the amount. We give you 10 days to respond in writing to prove if our facts are wrong."
			  Text 305, 115, 225, 40, "We will ask you to show proof of income, assets, or health insurance you did not report or that we could not verify. You may need to give us permission to check the facts with the source of data. We will tell you what happens if you do not sign for permission or do not help us."

			  Text 305, 155, 225, 10, "What is the Work Reporting System?"
			  Text 305, 165, 225, 40, "Minnesota employers must tell us when they hire someone. This information is used by the Child Support Program. We also use this information to see if a new employee is getting help from any of the programs listed above."
			  Text 305, 205, 225, 10, "How do we use it?"
			  Text 305, 215, 225, 40, "If the employee is getting help from any of these programs, the county worker gets a notice. If the client did not report the new job, the county worker will contact the client. The county worker may ask the client to show proof about the job. The client may need to give the county permission to check the facts with the employer. If a client does not help us check the information, they will lose benefits."
			  Text 305, 265, 225, 10, "The law limits who gets facts about you"
			  Text 305, 275, 225, 50, "The law limits the facts about you that we get from other agencies and the facts we give them. Contracts with the Minnesota Department of Human Services and those agencies also protect you. Only those agencies, the state, and the county agency where you apply for and get program benefits can use the facts about you. No one else can get the facts about you without your written permission."
			  ButtonGroup ButtonPressed
			    PushButton 15, 330, 100, 13, "Open DHS 2759", open_IEVS_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed IEVS Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_ievs_info_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_IEVS_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2759-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

'APPEAL RIGHTS
If left(confirm_appeal_rights_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Appeal Rights Discussed"+chr(9)+"No, I could not complete this", confirm_appeal_rights_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Appeal Rights"
			  Text 15, 25, 505, 10, "Appeal rights. An appeal is a legal process where a human services judge reviews a decision made by the agency. You may appeal a decision if:"
			  Text 20, 35, 500, 10, "You feel the agency did not act on your request for assistance."
			  Text 20, 45, 500, 10, "You do not agree with the action taken."
			  Text 15, 55, 505, 10, "You may represent yourself at the hearing, or you may have someone (an attorney, relative, friend or another person) speak for you."

			  Text 20, 65, 500, 20, "For emergency help, when your case is about an emergency and you need a faster decision on your appeal, you can ask for an emergency hearing in your appeal request. You can also request it by calling the Department of Human Services Appeals Division."
			  Text 20, 85, 500, 40, "For cash, child care and health care, you may appeal within 30 days from the date you received this notice by sending a written appeal request saying you do not agree with the decision. You can send this letter to the agency, or directly to the Appeals Division. If you show good cause for not appealing your cash, child care and health care within 30 days, the agency can accept your appeal for up to 90 days from the date of the notice. Good cause is when you have a good reason for not appealing on time. The Appeals Division will decide if your reason is a good cause reason. You can ask to meet informally with agency staff to try to solve the problem, but this meeting will not delay or replace your right to an appeal."
			  Text 20, 125, 500, 10, "For the Supplemental Nutrition Assistance Program, you may appeal within 90 days by writing or calling the agency or the Appeals Division."
			  Text 20, 135, 500, 10, "Submit your appeal request:"
			  Text 25, 145, 495, 10, "Online: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-0033-ENG"
			  Text 25, 155, 495, 10, "Write: Minnesota Department of Human Services Appeals Division P.O. Box 64941 St. Paul, MN 55164-0941"
			  Text 25, 165, 495, 10, "Fax: 651-431-7523"
			  Text 25, 175, 495, 10, "Call: Metro: 651-431-3600 Greater Minnesota: 800-657-3510 "
			  Text 20, 185, 500, 40, "If you want to keep receiving your benefits until the hearing, you must appeal within 10 days of the date on the agency’s notice of action letter or before the proposed action takes place in order to keep benefits in place. For most programs, if you file your appeal on time, you will get your benefits until the Appeals Division decides your appeal. If you lose your appeal, you may have to pay back the benefits you got while your appeal was pending. You can ask the agency to end your benefits until the decision. If you end your benefits and then win your appeal, you will be paid back for benefits that you should have received or, for child care assistance, your provider will be reimbursed for eligible costs that you paid or incurred. Ask your agency worker to explain how the timing of your appeal could affect your present or future assistance."
			  Text 15, 235, 505, 10, "You have the right to reapply at any time if your benefits stop."
			  Text 15, 245, 505, 20, "Access to free legal services. You may be able to get legal advice or help with an appeal from your local legal aid office. To find your local legal aid office, visit www.LawHelpMN.org or call 888-354-5522."

			  ButtonGroup ButtonPressed
			    PushButton 15, 265, 100, 13, "Open DHS 3353", open_appeal_rights_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Appeal Rights:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_appeal_rights_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_appeal_rights_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3353-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

'CIVIL RIGHTS NOTICE AND COMPLAINTS
If left(confirm_civil_rights_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Civil Rights Discussed"+chr(9)+"No, I could not complete this", confirm_civil_rights_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Civil Rights Notice and Complaints"
			  Text 15, 25, 505, 10, "Discrimination is against the law. The Minnesota Department of Human Services (DHS) does not discriminate on the basis of any of the following:"
			  Text 20, 35, 505, 10, "- race   - national origin   - religion   - public assistance status   - age   - sex   - color   - creed   - sexual orientation   - marital status   - disability   - political beliefs"

			  Text 15, 50, 275, 10, "Civil Rights Complaints"
			  Text 20, 60, 275, 20, "You have the right to file a discrimination complaint if you believe you were treated in a discriminatory way by a human services agency."
			  Text 20, 80, 275, 10, "Contact DHS directly only if you have a discrimination complaint:"
			  Text 25, 90, 275, 10, "Civil Rights Coordinator"
			  Text 25, 100, 275, 10, "Minnesota Department of Human Services"
			  Text 25, 110, 275, 10, "Equal Opportunity and Access Division"
			  Text 25, 120, 275, 10, "P.O. Box 64997 St. Paul, MN 55164-0997"
			  Text 25, 130, 275, 10, "651-431-3040 (voice) or use your preferred relay service"

			  Text 15, 140, 275, 10, "Minnesota Department of Human Rights (MDHR)"
			  Text 20, 150, 275, 20, "In Minnesota, you have the right to file a complaint with the MDHR if you believe you have been discriminated against because of any of the following:"
			  Text 25, 170, 275, 10, "- race   - sex   - color   - sexual orientation   - national origin   - marital status"
			  Text 25, 180, 275, 10, "- religion   - public assistance status   - creed   - disability"
			  Text 20, 190, 275, 10, "Contact the MDHR directly to file a complaint:"
			  Text 25, 200, 275, 10, "Minnesota Department of Human Rights"
			  Text 25, 210, 275, 10, "Freeman Building, 625 North Robert Street St. Paul, MN 55155"
			  Text 25, 220, 275, 10, "651-539-1100 (voice) 1-800-657-3704 (toll free) 651-296-9042 (fax)"
			  Text 25, 230, 275, 10, "Info.MDHR@state.mn.us (email)"


			  Text 15, 240, 275, 10, "U.S. Department of Health and Human Services' Office for Civil Rights (OCR)"
			  Text 20, 250, 275, 20, "You have the right to file a complaint with the OCR, a federal agency, if you believe you have been discriminated against because of any of the following:"
			  Text 25, 270, 275, 10, "- race   - age   - religion   - color   - disability   - national origin   - sex"
			  Text 20, 280, 275, 10, "Contact the OCR directly to file a complaint:"
			  Text 25, 290, 275, 10, "Director, U.S. Department of Health and Human Services' Office for Civil Rights"
			  Text 25, 300, 275, 10, "200 Independence Avenue SW, Room 509F HHH Building Washington, DC 20201"
			  Text 25, 310, 275, 10, "1-800-368-1019 (voice)  1-800-537-7697 (TDD)"
			  Text 25, 320, 275, 10, "Complaint Portal: https://ocrportal.hhs.gov/ocr/portal/lobby.jsf"

			  Text 305, 55, 225, 60, "In accordance with Federal civil rights law and U.S. Department of Agriculture (USDA) civil rights regulations and policies, the USDA, its Agencies, offices, and employees, and institutions participating in or administering USDA programs are prohibited from discriminating based on race, color, national origin, sex, religious creed, disability, age, political beliefs, or reprisal or retaliation for prior civil rights activity in any program or activity conducted or funded by USDA."
			  Text 305, 115, 225, 70, "Persons with disabilities who require alternative means of communication for program information (e.g. Braille, large print, audiotape, American Sign Language, etc.), should contact the Agency (State or local) where they applied for benefits. Individuals who are deaf, hard of hearing or have speech disabilities may contact USDA through the Federal Relay Service at 1-800-877-8339. Additionally, program information may be made available in languages other than English."
			  Text 305, 185, 225, 60, "To file a program complaint of discrimination, complete the USDA Program Discrimination Complaint Form, (AD-3027) found online at: http://www.ascr.usda.gov/complaint_filing_cust.html, and at any USDA office, or write a letter addressed to USDA and provide in the letter all of the information requested in the form. To request a copy of the complaint form, call 1-866- 632-9992. Submit your completed form or letter to USDA by:"
			  Text 310, 245, 225, 10, "(1) mail: U.S. Department of Agriculture"
			  Text 315, 255, 225, 10, "Office of the Assistant Secretary for Civil Rights"
			  Text 315, 265, 225, 10, "1400 Independence Avenue, SW"
			  Text 315, 275, 225, 10, "Washington, DC 20250-9410;"
			  Text 310, 285, 225, 10, "(2) fax: 202-690-7442; or"
			  Text 310, 295, 225, 10, "(3) email: program.intake@usda.gov"
			  Text 310, 305, 225, 10, "This institution is an equal opportunity provider."

			  ButtonGroup ButtonPressed
			    PushButton 15, 340, 100, 13, "Open DHS 3353", open_civil_rights_rights_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Civil Rights Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_civil_rights_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_civil_rights_rights_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3353-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

' 'COVER LETTER
' Do
' 	Do
' 		err_msg = ""
'
' 		Dialog1 = ""
' 		BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
' 		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Cover Letter Discussed"+chr(9)+"No, I could not complete this", confirm_cover_letter_read
' 		  ButtonGroup ButtonPressed
' 		    PushButton 465, 365, 80, 15, "Continue", continue_btn
' 		  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
' 		  GroupBox 10, 15, 530, 345, "Hennepin County Cover Letter"
'
'
'
' 		  Text 10, 370, 210, 10, "Confirm you have reviewed Hennepin County Information Information:"
' 		EndDialog
'
' 		dialog Dialog1
'
'
' 		cancel_confirmation
' 	Loop until err_msg = ""
' 	Call check_for_password(are_we_passworded_out)
' Loop until are_we_passworded_out = FALSE
' save_your_work

'PROGRAM INFORMATION FOR CASH, FOOD, CHILD CARE - 2920
If left(confirm_program_information_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Program Information Discussed"+chr(9)+"No, I could not complete this", confirm_program_information_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Program Information for cash, food, and child care programs"
			  Text 15, 25, 505, 10, "How do you apply for help?"
			  Text 20, 35, 505, 10, "If you do not have enough money to meet your basic needs, you can apply to find out if you are eligible for these assistance programs."
			  Text 25, 45, 505, 10, "Apply online at MNbenefits.mn.gov."
			  Text 25, 55, 505, 10, "Mail or bring your completed application to your county human services agency"
			  Text 20, 65, 505, 10, "Food and cash programs require an interview with a worker. Most of the time this can be a phone interview. You will need to bring proof of:"
			  Text 20, 75, 505, 10, "- Who you are   - Where you live   - What family members live with you   - What your income is   - What you own."
			  Text 20, 85, 505, 10, "Whether or not you can receive help and how much you receive may depend on:"
			  Text 20, 95, 505, 10, "- How long you have lived in Minnesota   - How many people live with you   - How much income you and these people receive each month."
			  Text 20, 105, 505, 10, "Each program has different rules."

			  Text 15, 115, 275, 20, "Cash assistance is provided to help you meet your basic needs, if you are eligible. Some of the programs have time limits. Cash programs include:"
			  Text 20, 135, 275, 10, "Diversionary Work Program (DWP)"
			  Text 25, 145, 275, 30, "A short-term work program that provides employment services and basic living costs to eligible families. DWP is for families who are working or looking for work, but need help with basic living expenses and have not MFIP or DWP in the last 12 months."
			  Text 20, 175, 275, 10, "Minnesota Family Investment Program (MFIP)"
			  Text 25, 185, 275, 20, "A monthly cash assistance program for families with children under 19 or pregnant women, and who have low incomes."
			  Text 20, 205, 275, 10, "General Assistance (GA)"
			  Text 25, 215, 275, 10, "A monthly cash payment for adults who are unable to work who:"
			  Text 30, 225, 275, 10, "- Have little or no income and will soon return to work, or"
			  Text 30, 235, 275, 10, "- Are waiting to get help from other state or federal programs."
			  Text 20, 245, 275, 10, "Minnesota Supplemental Aid (MSA)"
			  Text 25, 255, 275, 10, "A small extra monthly cash payment for adults who are eligible for federal SSI."
			  Text 20, 265, 275, 10, "Group Residential Housing (GRH)"
			  Text 25, 275, 275, 20, "A monthly payment that helps pay room and board costs for people who live in authorized settings and are:"
			  Text 130, 285, 275, 10, "- Age 65 or older "
			  Text 130, 295, 275, 10, "- Disabled and age 18 or older, or "
			  Text 130, 305, 275, 10, "- Have blindness."
			  Text 20, 315, 275, 10, "Refugee Cash Assistance (RCA)"
			  Text 25, 325, 275, 10, "A monthly cash payment for refugees and asylees. RCA is for people who:"
			  Text 30, 335, 275, 10, "- Have been in the United States eight months or less, and "
			  Text 30, 345, 275, 10, "- Have refugee or asylee status."

			  Text 305, 115, 225, 20, "Minnesota's Child Care Assistance Program makes quality child care affordable for families with low incomes, from the following programs:"
			  Text 310, 135, 225, 10, "MFIP Child Care"
			  Text 315, 145, 225, 30, "Families who receive assistance from the Diversionary Work Program or Minnesota Family Investment Program are eligible for child care if the parents are in work related activities."
			  Text 310, 175, 225, 10, "Transition Year Child Care"
			  Text 315, 185, 225, 30, "Available to families for up to 12 consecutive months after their Diversionary Work Program or Minnesota Family Investment Program case closes."
			  Text 310, 215, 225, 10, "Basic Sliding Fee Child Care"
			  Text 315, 225, 225, 10, "Available for other families with low incomes."
			  Text 310, 240, 225, 10, "Supplemental Nutrition Assistance Program (SNAP)"
			  Text 315, 250, 225, 30, "A federal program that helps Minnesotans with low income buy food. Benefits are available through EBT cards that can be used like money. Benefits are for:"
			  Text 320, 275, 225, 10, "- Single people"
			  Text 320, 285, 225, 10, "- Families with or without children"
			  Text 315, 295, 225, 20, "Your income, the size of your household, and your housing costs determines how much you can receive."


			  ButtonGroup ButtonPressed
			    PushButton 405, 340, 100, 13, "Open DHS 2920", open_program_info_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Program Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_program_information_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_program_info_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2920-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

'DOMESTIC VIOLENCE INFORMATION - 3477
If left(confirm_DV_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Domestic Violence Discussed"+chr(9)+"No, I could not complete this", confirm_DV_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Domestic Violence Information"
			  Text 15, 25, 505, 10, "If you are in danger from domestic violence or abuse and need help, call:"
			  Text 20, 35, 505, 10, "The National Domestic Violence Hotline at 800-799-7233, (TTY:800-787-3224)"
			  Text 20, 45, 505, 10, "The Minnesota Coalition for Battered Women at 866-289-6177"
			  Text 20, 55, 505, 10, "The Minnesota Day One Emergency Shelter and Crisis Hotline at 800-223-1111"

			  Text 15, 65, 275, 10, "What is domestic violence?"
			  Text 20, 75, 275, 40, "Domestic violence or abuse is what someone says or does over and over again to make you feel afraid or to control you. People who are elderly, frail, have a disability, or who depend on others for assistance may not be able to protect themselves from domestic violence or abuse. Minnesota has a law to protect and assist people who are vulnerable to abuse or who are not able to care for themselves. Examples of violence or abuse include:"
			  Text 25, 115, 275, 10, "- Swearing or screaming at you"
			  Text 25, 125, 275, 10, "- Calling you names"
			  Text 25, 135, 275, 10, "- Taking money or property without permission"
			  Text 25, 145, 275, 10, "- Threatening to hurt you or others you care about"
			  Text 25, 155, 275, 10, "- Failing to provide care for you"
			  Text 25, 165, 275, 10, "- Not letting you leave your house"
			  Text 25, 175, 275, 10, "- Blaming you for everything that goes wrong"
			  Text 25, 185, 275, 10, "- Stalking you"
			  Text 25, 195, 275, 10, "- Being touched against your wishes or forced to have sex"
			  Text 25, 205, 275, 10, "- Choking, grabbing, hitting, pushing, pinching or kicking you."

			  Text 15, 215, 275, 20, "What services are available to victims of domestic violence or abuse?"

			  Text 20, 225, 275, 10, "Toll-free Hotlines have counselors who provide services, including:"
			  Text 25, 235, 275, 10, "- Crisis counseling"
			  Text 25, 245, 275, 10, "- Safety planning"
			  Text 25, 255, 275, 10, "- Assistance with finding shelter."
			  Text 20, 265, 275, 10, "Referrals to other organizations including:"
			  Text 25, 275, 275, 10, "- Legal services support groups"
			  Text 25, 285, 275, 10, "- Advocacy with the police."


			  Text 305, 65, 225, 10, "Safe At Home (SAH) Program"
			  Text 310, 75, 225, 60, "The Safe At Home (SAH) Program is a statewide address confidentiality program that assists survivors of domestic violence, sexual assault, stalking and others who fear for their safety by providing a substitute address for people who move or are about to move to a new location unknown to their aggressors. For information on this program, contact Safe At Home at 651-201-1399 or 866-723-3035."
			  Text 305, 135, 225, 10, "Vulnerable adults"
			  Text 310, 145, 225, 30, "Call the Senior LinkAge Line at 800-333-2433 to report concerns and to help a vulnerable adult get needed protection and assistance. Ask your worker for more resource information."
			  Text 305, 175, 225, 10, "What are domestic violence waivers?"
			  Text 310, 185, 225, 30, "If you are eligible for public assistance and you experience domestic violence, certain program requirements may not apply in your situation."
			  Text 310, 215, 225, 30, "If domestic violence or abuse makes it hard for you to follow program rules, talk to your county worker."


			  ButtonGroup ButtonPressed
			    PushButton 15, 340, 100, 13, "Open DHS 3477", open_DV_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Domestic Violence Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_DV_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_DV_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

'DO YOU HAVE A DISABILITY - 4133
If left(confirm_disa_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Disability Information Discussed"+chr(9)+"No, I could not complete this", confirm_disa_read
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
			  GroupBox 10, 15, 530, 345, "Do you have a Disability"
			  Text 15, 25, 505, 10, "Please tell us if you have a disability so we can help you access human services programs and benefits."

			  Text 15, 35, 275, 10, "What medical conditions may be disabilities?"
			  Text 20, 45, 275, 20, "A disability is a physical, sensory, or mental impairment that materially limits a major life activity. Types of disabilities may include:"
			  Text 25, 65, 275, 10, "- Diseases like diabetes, epilepsy or cancer"
			  Text 25, 75, 275, 10, "- Learning disorders like dyslexia"
			  Text 25, 85, 275, 10, "- Developmental delays"
			  Text 25, 95, 275, 10, "- Clinical depression"
			  Text 25, 105, 275, 10, "- Hearing loss or low vision"
			  Text 25, 115, 275, 10, "- Movement restrictions like trouble with walking, reaching or grasping"
			  Text 25, 125, 275, 10, "- History of alcohol or drug addiction"
			  Text 35, 135, 275, 10, "(current illegal drug use is not a disability)"
			  Text 20, 145, 275, 30, "If you are asking for or are getting benefits through either a county human services agency or the Minnesota Department of Human Services, that office will let you know if you have a disability using information from you and your doctor."

			  Text 305, 35, 225, 10, "What help is available?"
			  Text 310, 45, 225, 20, "If you have a disability, your county or the state human services agency can help you by:"
			  Text 315, 65, 225, 20, "- Calling you or meeting with you in another place if you are not able to come into the office"
			  Text 315, 85, 225, 10, "- Using a sign language interpreter"
			  Text 315, 95, 225, 20, "- Giving you letters and forms in other formats like computer files, audio recordings, large print or Braille"
			  Text 315, 115, 225, 10, "- Telling you the meaning of the information we give you"
			  Text 315, 125, 225, 10, "- Helping you fill out forms"
			  Text 315, 135, 225, 10, "- Helping you make a plan so you can work even with your disability"
			  Text 315, 145, 225, 10, "- Sending you to other services that may help you"
			  Text 315, 155, 225, 20, "- Helping you to appeal agency decisions about you if you disagree with them"
			  Text 310, 175, 225, 30, "You will not have to pay extra for help. If you want help, ask your agency as soon as possible. An agency may not be able to accommodate requests made within 48 hours of need."

			  Text 15, 205, 505, 10, "How does the law protect people with disabilities?"
			  Text 20, 215, 505, 40, "The Americans with Disabilities Act (ADA) and the ADA Amendments Act are federal laws, and the Minnesota Human Rights Act is a state law. Each gives individuals with disabilities the same legal rights and protections as people without disabilities, including access to public assistance benefits. You will not be denied benefits because you have a disability. Your benefits will not be stopped because of your disability. If your disability makes getting benefits hard for you, your county human services agency will help you access all of the programs that are available to you."

			  ButtonGroup ButtonPressed
			    PushButton 15, 340, 100, 13, "Open DHS 4133", open_disa_doc
			  Text 10, 370, 210, 10, "Confirm you have reviewed Disability Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_disa_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = open_disa_doc Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-4133-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work



'MFIP CASES
	'Reporting Responsibilities for MFIP Households (DHS-2647) (PDF).
	'Notice of Requirement to Attend MFIP Overview (DHS-2929) (PDF). See 0028.09 (ES Overview/SNAP E&T Orientation).
	'Family Violence Referral (DHS-3323) (PDF) and
If family_cash_case_yn = "Yes" Then
	If left(confirm_mfip_forms_read, 4) <> "YES!" Then
		Do
			Do
				err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"

				  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Forms Discussed"+chr(9)+"No, I could not complete this", confirm_mfip_forms_read
				  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
				  GroupBox 10, 15, 530, 345, "MFIP Cases"

				  GroupBox 10, 25, 530, 105, "Reporting Responsibilities for MFIP Households (DHS-2647)"
				  Text 15, 40, 505, 10, "Changes you must report: Anything that could impact eligibility. Particularly: Income, Assets, Household Comp"
				  Text 15, 50, 505, 10, "When do Changes need to be Reported: On the monthly Household Report Form, if you do not have one, within 10 days of the change"
				  Text 15, 60, 505, 10, "How to report Changes: On any Report Form or call the county."
				  GroupBox 10, 125, 530, 105, "Notice of Requirement to Attend MFIP Overview (DHS-2929)"
				  Text 15, 140, 505, 10, "All MFIP caregivers are required to attend an MFIP overview and participate in Employment Services."
				  Text 15, 150, 505, 10, "If you do not go to your scheduled overview meeting without good reason, your MFIP grant may be reduced until you go to the meeting."
				  Text 15, 160, 505, 10, "Call the contact person above if you: - Need child care or help getting to the meeting - Have problems attending the meeting."
				  GroupBox 10, 225, 530, 105, "Family Violence Referral (DHS-3323)"
				  Text 15, 240, 505, 10, "If you, or someone in your home is a victim of domestic abuse the county can help you."
				  Text 15, 250, 505, 10, "You can also call the National Domestic Violence Hot Line at (800) 799-7233 or Legal Aid at (888) 354-5522."
				  Text 15, 260, 505, 20, "Some of the Minnesota Family Investment Program (MFIP) rules do not apply to domestic abuse victims. You must tell us about the abuse and have a special employment plan that includes activities to help keep your family safe. Please talk to your worker or an advocate if you want to know about this."
				  ButtonGroup ButtonPressed
				    PushButton 465, 365, 80, 15, "Continue", continue_btn
				    PushButton 430, 22, 100, 13, "Open DHS 2647", open_cs_2647_doc
					PushButton 430, 122, 100, 13, "Open DHS 2929", open_cs_2929_doc
					PushButton 430, 222, 100, 13, "Open DHS 3323", open_cs_3323_doc
				  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Specific Information:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If confirm_mfip_forms_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

				If ButtonPressed = open_cs_2647_doc OR ButtonPressed = open_cs_2929_doc OR ButtonPressed = open_cs_3323_doc Then
					err_msg = "LOOP"
					If ButtonPressed = open_cs_2647_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2647-ENG"
					If ButtonPressed = open_cs_2929_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2929-ENG"
					If ButtonPressed = open_cs_3323_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3323-ENG"
				End If

				IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE
	End If
	save_your_work

	If absent_parent_yn = "Yes" Then
		'In cases where there is at least 1 non-custodial parent:
			'Understanding Child Support - A Handbook for Parents (DHS-3393) (PDF).
			'Referral to Support and Collections (DHS-3163B) (PDF). (This is in addition to the Combined Application Form, for EACH non-custodial parent). See 0012.21.03 (Support From Non-Custodial Parents).
			'Cooperation with Child Support Enforcement (DHS-2338) (PDF). See 0012.21.06 (Child Support Good Cause Exemptions).
		'If a non-parental caregiver applies,
			'MFIP Child Only Assistance (DHS-5561) (PDF).
		If left(confirm_mfip_cs_read, 4) <> "YES!" Then
			Do
				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
					  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Child Support Discussed"+chr(9)+"No, I could not complete this", confirm_mfip_cs_read
					  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
					  GroupBox 10, 15, 530, 345, "MFIP Case with at least 1 ABPS - Child Support Information"

					  GroupBox 10, 25, 530, 105, "Understanding Child Support - A Handbook for Parents (DHS-3393)"
					  Text 15, 40, 505, 20, "Every child needs financial and emotional support. Every child has the right to this support from both parents. Devoted parents can be loving and supportive forces in a child's life. Even when parents do not live together, they need to work together to support their child."
					  Text 15, 60, 505, 20, "Minnesoda Child Support and Hennepin County Child Support provide support and guidance. The Handbook 'Understanding Child Support' provides information about the details of these programs."
					  GroupBox 10, 125, 530, 105, "Referral to Support and Collections (DHS-3163B)"
					  Text 15, 140, 505, 10, "Purpose of form: The child support agency will use the information you give to help collect support."
					  Text 15, 150, 505, 20, "How to complete this form: Fill in each blank. If there are boxes, check the box or boxes that fit your situation. Complete a separate form for each parent or alleged parent other than yourself."
					  Text 15, 170, 505, 20, "Please read the booklet 'Understanding Child Support: A Handbook for Parents' (DHS-3393) before signing. The booklet explains information about the child support services you may be receiving."
					  GroupBox 10, 225, 530, 105, "Cooperation with Child Support Enforcement (DHS-2338)"
					  Text 15, 240, 505, 10, "This notice explains your rights and responsibilities for cooperating with the MN Department of Human Services, Child Support Division."
					  Text 15, 250, 505, 10, "Cooperation with the child support agency includes answering questions, filling out forms, and appearing at appointments and/or court hearings."
					  Text 15, 260, 505, 40, "This notice also explains how you make a 'good cause claim' that gives you the right not to cooperate if your claim is granted. If you choose to claim good cause and your county child support agency is currently collecting your child support payments, the county will immediately stop collecting those payments for the child(ren) you name on the attached form. The county will stop providing all child support services until it makes a decision on your good cause claim. If you are granted a good cause exemption, the child support agency will close your case."
					  If relative_caregiver_yn = "Yes" Then
						  GroupBox 10, 325, 530, 40, "If Non-Custodial Caregiver - MFIP Child Only Assistance (DHS-5561)"
						  Text 15, 335, 505, 30, "The Minnesota Department of Human Services has assistance programs available to help children who are cared for and supported by their relatives. This brochure answers some frequently asked questions relatives may have about the Minnesota Family Investment Program (MFIP)"
					  End If
					  ButtonGroup ButtonPressed
					    PushButton 465, 365, 80, 15, "Continue", continue_btn
					    PushButton 430, 22, 100, 13, "Open DHS 3393", open_cs_3393_doc
						PushButton 430, 122, 100, 13, "Open DHS 3163B", open_cs_3163B_doc
						PushButton 430, 222, 100, 13, "Open DHS 2338", open_cs_2338_doc
						If relative_caregiver_yn = "Yes" Then PushButton 430, 322, 100, 13, "Open DHS 5561", open_cs_5561_doc

					  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Child Support Information:"
					EndDialog

					dialog Dialog1
					cancel_confirmation

					If confirm_mfip_cs_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

					If ButtonPressed = open_cs_3393_doc OR ButtonPressed = open_cs_3163B_doc OR ButtonPressed = open_cs_2338_doc OR ButtonPressed = open_cs_5561_doc Then
						err_msg = "LOOP"
						If ButtonPressed = open_cs_3393_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3393-ENG"
						If ButtonPressed = open_cs_3163B_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3163B-ENG"
						If ButtonPressed = open_cs_2338_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2338-ENG"
						If ButtonPressed = open_cs_5561_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-5561-ENG"
					End If

					IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				Loop until err_msg = ""
				Call check_for_password(are_we_passworded_out)
			Loop until are_we_passworded_out = FALSE
		End If
		save_your_work
	End If

	If left(minor_caregiver_yn, 3) = "Yes" Then
		'If there is a custodial parent under 20, the
			'Notice of Requirement to Attend School (DHS-2961) (PDF) and
			'Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS-2887) (PDF).
		'If there is a custodial parent under age 18, the
			'MFIP for Minor Caregivers (DHS-3238) (PDF) brochure.
		If left(confirm_minor_mfip_read, 4) <> "YES!" Then
			Do
				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
					  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Minor Caregiver Discussed"+chr(9)+"No, I could not complete this", confirm_minor_mfip_read
					  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
					  GroupBox 10, 15, 530, 345, "MFIP Case Minor Caregiver Cases"

					  GroupBox 10, 25, 530, 105, "Notice of Requirement to Attend School (DHS-2961)"
					  Text 15, 40, 505, 10, "This form tells you that, unless you are exempt, you must attend school and what will happen if you do not go to school."
					  Text 15, 50, 505, 20, "The first step is for us to complete an assessment with you. We will review your educational progress, needs, literacy level, family circumstances, skills, and work experience. We will see if you need child care or other services so you can go to school."
					  Text 15, 70, 505, 10, "If you do not cooperate or do not attend school, without good cause, we will send you a notice. This notice will tell you that your MFIP grant may be reduced. "
					  GroupBox 10, 125, 530, 105, "Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS-2887)"
					  Text 15, 140, 505, 20, "If you are a teen parent under the age of 20, and do not have a high school diploma or an equivalent, you are expected to attend an approved educational program to qualify for the Minnesota Family Investment Program."
					  Text 15, 160, 505, 20, "Earning your diploma is the first step in getting ready for a job. County human services staff will help you with counseling, child care, and transportation so you can go to school. They will also help you find a school program that is best for you."
					  Text 15, 180, 505, 10, "If you fail to attend school, without good cause, your human services worker will reduce your grant by 10 percent or more of your standard of need."
					  If minor_caregiver_yn = "Yes - Caregiver is under 18" Then
						  GroupBox 10, 225, 530, 105, "MFIP for Minor Caregivers (DHS-3238)"
						  Text 15, 240, 505, 20, "You are a minor caregiver if: "
						  Text 25, 250, 505, 20, "- You are younger than 18 - You have never been married - You are not emancipated and - You are the parent of a child(ren) living in the same household."
						  Text 15, 260, 505, 10, "If you are a minor caregiver, to receive benefits and services, you must be living: "
						  Text 25, 270, 505, 20, "- With a parent or with an adult relative caregiver or with a legal guardian or - In an agency-approved living arrangement."
						  Text 15, 280, 505, 10, "A social worker must approve any exception(s) to your living arrangement."
					  End If
					  ButtonGroup ButtonPressed
					  	PushButton 465, 365, 80, 15, "Continue", continue_btn
					    PushButton 430, 22, 100, 13, "Open DHS 2961", open_cs_2961_doc
						PushButton 430, 122, 100, 13, "Open DHS 2887", open_cs_2887_doc
						If minor_caregiver_yn = "Yes - Caregiver is under 18" Then PushButton 430, 222, 100, 13, "Open DHS 3238", open_cs_3238_doc

					  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Minor Caregiver Information:"
					EndDialog

					dialog Dialog1
					cancel_confirmation

					If confirm_minor_mfip_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

					If ButtonPressed = open_cs_2961_doc OR ButtonPressed = open_cs_2887_doc OR ButtonPressed = open_cs_3238_doc Then
						err_msg = "LOOP"
						If ButtonPressed = open_cs_2961_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Legacy/DHS-2961-ENG"
						If ButtonPressed = open_cs_2887_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2887-ENG"
						If ButtonPressed = open_cs_3238_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3238-ENG"
					End If

					IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
				Loop until err_msg = ""
				Call check_for_password(are_we_passworded_out)
			Loop until are_we_passworded_out = FALSE
		End If
		save_your_work
	End If
End If


If snap_case = True OR pend_snap_on_case = "Yes" OR mfip_status <> "INACTIVE" Then
	'SNAP CASES'
		'Supplemental Nutrition Assistance Program reporting responsibilities (DHS-2625).
		'Facts on Voluntarily Quitting Your Job If You Are on the Supplemental Nutrition Assistance Program (SNAP) (DHS-2707).
		'Work Registration Notice (DHS-7635).
	If left(confirm_snap_forms_read, 4) <> "YES!" Then
		Do
			Do
				err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
				  DropListBox 90, 40, 105, 45, "Select One..."+chr(9)+"Six-Month"+chr(9)+"Change"+chr(9)+"Monthly", snap_reporting_type
				  EditBox 115, 55, 40, 15, next_revw_month
				  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! SNAP Forms Discussed"+chr(9)+"No, I could not complete this", confirm_snap_forms_read
				  Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
				  GroupBox 10, 15, 530, 345, "SNAP Case"
				  GroupBox 10, 25, 530, 105, "Supplemental Nutrition Assistance Program reporting responsibilities (DHS-2625)"
				  Text 15, 45, 75, 10, "This case is subject to "
				  Text 200, 45, 40, 10, "reporting."
				  Text 15, 60, 95, 10, "Your next renewal will be for "
				  Text 160, 60, 310, 10, ". Which means you will need to complete the required form and process in the month before."
				  Text 15, 75, 185, 10, "Explain reporting details based on the reporter type."
				  Text 15, 100, 395, 10, "Timely reporting of changes means the change is reported by the 10th of the month following the month of the change."
				  GroupBox 10, 125, 530, 105, "Facts on Voluntarily Quitting Your Job If You Are on SNAP (DHS-2707)"
				  Text 15, 140, 505, 10, "If you or someone else in your household has a job and quits without a good reason, your household might not get SNAP benefits."
				  Text 15, 150, 505, 20, "The penalty does not apply if the person who quit a job: "
				  Text 25, 160, 505, 20, "- Was fired, or forced to leave the job, or had hours cut back by the employer - Was self-employed - Left a job that was less than 30 hours per week"
				  Text 15, 170, 505, 10, "The penalty also does not apply if you can prove the person had 'good reason' to quit the job. The form has some examples of 'good reasons'."
				  GroupBox 10, 225, 530, 105, "Work Registration Notice (DHS-7635)"
				  Text 15, 240, 505, 10, "In order to be eligible for benefits you must cooperate in any efforts regarding work registration. "
				  Text 15, 250, 505, 10, "If you do not follow any of the work requirements listed above your benefits may end."
				  ButtonGroup ButtonPressed
				    PushButton 465, 365, 80, 15, "Continue", continue_btn
				    PushButton 430, 20, 100, 15, "Open DHS 2625", open_cs_2625_doc
					PushButton 25, 85, 90, 13, "Six Month Reporting", explain_six_month_rept
					PushButton 115, 85, 90, 13, "Change Reporting", explain_change_rept
					PushButton 205, 85, 90, 13, "Monthly Reporting", explain_monthly_rept
				    PushButton 430, 120, 100, 15, "Open DHS 2707", open_cs_2707_doc
				    PushButton 430, 220, 100, 15, "Open DHS 7635", open_cs_7635_doc
				  Text 10, 370, 210, 10, "Confirm you have reviewed SNAP Specific Information:"
				EndDialog

				dialog Dialog1
				cancel_confirmation

				If confirm_snap_forms_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."
				If confirm_snap_forms_read = "YES! SNAP Forms Discussed" Then
					If snap_reporting_type = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since you have reviewed SNAP information, select the correct reporting type for this case to ensure the best information is provided to the household."
					If Trim(next_revw_month) = "" Then err_msg = err_msg & vbNewLine & "* Since you have reviewed SNAP information, indicate the next review month for this case."
				End If

				If ButtonPressed = open_cs_2625_doc OR ButtonPressed = open_cs_2707_doc OR ButtonPressed = open_cs_7635_doc Then
					err_msg = "LOOP"
					If ButtonPressed = open_cs_2625_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2625-ENG"
					If ButtonPressed = open_cs_2707_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2707-ENG"
					If ButtonPressed = open_cs_7635_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-7635-ENG"
				End If

				If ButtonPressed = explain_six_month_rept OR ButtonPressed = explain_change_rept OR ButtonPressed = explain_monthly_rept Then
					err_msg = "LOOP"
					If ButtonPressed = explain_six_month_rept Then MsgBox "SIX MONTH REPORTING" & vbCr & vbCR &_
																		  "There are only TWO changes that are required to be reported:" & vbCr &_
																		  "  - Income received in any month exceeds 130% FPG for the Household "  & vbCr &_
																		  "    Size." & vbCr &_
																		  "  - For any ABAWD, a change in work or job activities that cause their "  & vbCr &_
																		  "    hours to fall below 20 hours per week, averaged 80 hours monthly." & vbCr & vbCr &_
																		  "It can be beneficial to report other changes, and we encourage you to do this. Examples include:" & vbCr &_
																		  "  - Address Changes " & vbCr &_
																		  "    (We communicate via mail and missing mail can cause your "  & vbCr &_
																		  "    benefits to close for lack of response.)" & vbCr &_
																		  "  - Decreases in Income " & vbCr &_
																		  "    (Income is used to determine your benefit amount and any reduction " & vbCr &_
																		  "     MAY cause your benefit amount to increase.)" & vbCr &_
																		  "  - Increases in Housing Expense or New Utility Responsibilities " & vbCr &_
																		  "    (Expenses are used to offset income and changes here could " & vbCr &_
																		  "    change your benefit amount.)" & vbCr &_
																		  "  - Other Expenses " & vbCr &_
																		  "    ie. Child Care, Child Support, sometimes Medical Expenses " & vbCr &_
																		  "    (These can also impact your benefit amount.)" & vbCr & vbCr &_
																		  "As a Six-Month Reporter, you are certified for six months at a time, which means you will have a review within six months."
					If ButtonPressed = explain_change_rept Then MsgBox "CHANGE REPORTING" & vbCr & vbCR &_
																	   "Changes that are required to be reported:" & vbCr &_
																	   "  - A change in the source of income, including starting or stopping a " & vbCr &_
																	   "    job, if the change in employment is accompanied by a change in " & vbCr &_
																	   "    income." & vbCr &_
																	   "  - A change in more than $125 per month in gross earned income." & vbCr &_
																	   "  - A change of more than $125 in the amount of unearned income, " & vbCr &_
																	   "    EXCEPT changes related to public assistance." & vbCr &_
																	   "  - A change in unit composition." & vbCr &_
																	   "  - A change in residence." & vbCr &_
																	   "  - A change in housing expense due to residency change." & vbCr &_
																	   "  - A change in legal obligation to pay child support." & vbCr &_
																	   "  - For any ABAWD, a change in work or job activities that cause their " & vbCr &_
																	   "    hours to fall below 20 hours per week, averaged 80 hours monthly." & vbCr & vbCr &_
																	   "As a Change Reporter, you typically have a certification period of a year but it could be two years."
					If ButtonPressed = explain_monthly_rept Then MsgBox "MONTHLY REPORTING" & vbCr & vbCR &_
																	    "Monthly reporters are required to submit a Household Report Form every month with income and change verifications attached." & vbCr & vbCr &_
																		"The Household Report Form must be answered in its entirety. Any unanswered question will make the form incomplete and ongoing benefits will not be able to be processed. The form includes all changes that must be reported." & vbCr & vbCr &_
																	    "As a Monthly Reporter, you are certified for twelve months at a time, which means you will have a review within twelve months." & vbCr &_
																		"However the system will close your benefits if the monthly Household Report Form is not received, processed, and all verifications attached."

				End If

				IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE
	End If
	save_your_work
End If
'Employment Services Registration.

'REPORTING

'Additional Important Information.

'Penalty Warnings.


' Call provide_resources_information(case_number_known, create_case_note, note_detail_array, allow_cancel)
Call provide_resources_information(True, False, note_detail_array, False)

If left(confirm_recap_read, 4) <> "YES!" Then
	Do
		Do
			err_msg = ""
			Call create_verifs_needed_list(verifs_selected, verifs_needed)
			If trim(verifs_needed) <> "" Then
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

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Recap Discussed"+chr(9)+"No, I could not complete this", confirm_recap_read
			  Text 200, 10, 335, 10, "The Interview Information has been completed. Review the information and next steps with the resident."
			  GroupBox 10, 20, 530, 340, "CASE INTERVIEW WRAP UP"

			  ' Text 15, 30, 505, 10, "What would be helpful here?"
			  y_pos = 45
			  If trim(verifs_needed) = "" Then
				  Text 15, 35, 505, 10, "THERE ARE NO REQUESTED VERIFICATIONS ENTERED INTO THE SCRIPT"
				  Text 15, 50, 505, 10, "Since there are no verifications requested, the program requests should be processed."
				  y_pos = 65
			  Else
				  If verif_req_form_sent_date = "" Then Text 15, 35, 505, 10, "Requested Verifications Entered Into the Script"
				  If verif_req_form_sent_date <> "" Then Text 15, 35, 505, 10, "Requested Verifications Entered Into the Script. Request sent on " & verif_req_form_sent_date
			  	  For each verif_item in verifs_array
					  If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
					  Text 25, y_pos, 505, 10, verif_item

					  verif_counter = verif_counter + 1
					  y_pos = y_pos + 10
				  Next
			  End If


			  ButtonGroup ButtonPressed
			  	PushButton 15, y_pos, 100, 15, "Update Verifications", verif_button
			  y_pos = y_pos + 20


			  ' Text 15, 30, 505, 10, "Programs being Requested/Renewed:"
			  ' Text 20, 40, 505, 10, "SNAP"
			  ' Text 20, 50, 505, 10, "Cash - MFIP"
			  ' Text 20, 60, 505, 10, "Housing Support - GRH"
			  ' Text 15, 75, 505, 10, "Next Steps:"
			  ' Text 20, 85, 505, 10, "We need verifications before we can make a determination on your case. Are you clear on what those are? You will also receive a notice in the mail."
			  ' Text 20, 95, 505, 10, "If you need an EBT Card - call or go in."
			  ' Text 20, 105, 505, 10, "I will be processing your case. "
			  ' Text 25, 115, 505, 10, "APPLICATION - the benefits are typically available the day after appproval. "
			  ' Text 25, 125, 505, 10, "RECERT - the benefits should be available on your regular day."
			  ' Text 20, 135, 505, 10, "Watch your mail for approval notices to see the benefit amount."
			  Text 15, y_pos, 505, 10, "Your address and phone number are our best way to contact you."
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "It is vital that you let us know if you address or phone number has changed"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "You may miss important requests or notices if we have an old address."
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "Our mail does not forward to address changes, so we need to know the correct address for you"
			  y_pos = y_pos + 15
			  Text 15, y_pos, 505, 10, "Please be sure to follow program rules and requirements"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "Failure to report changes and information timely can have negative impacts:"
			  y_pos = y_pos + 10
			  Text 25, y_pos, 505, 10, "End of benefits"
			  y_pos = y_pos + 10
			  Text 25, y_pos, 505, 10, "Overpayments"
			  y_pos = y_pos + 10
			  Text 25, y_pos, 505, 10, "Future ineligibility"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "We receive information from other sources about you and may impact your eligibility and benefit level."
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "If you are unsure of program rules and requirements, the forms we reviewed earlier can always be resent, or you can call us with questions."
			  y_pos = y_pos + 15
			  Text 15, y_pos, 505, 10, "Contact to Hennepin County"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "By Phone - 612-596-1300. The phone lines are open Monday - Friday 8:00 - 4:30"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "In person - Not Available Currently"
			  y_pos = y_pos + 10
			  Text 20, y_pos, 505, 10, "Online - MNbenefits or InfoKeep"
              y_pos = y_pos + 20

              Text 15, y_pos, 250, 10, "Summarize what is happening with this case:"
              y_pos = y_pos + 10
              EditBox 15, y_pos, 520, 15, case_summary

			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			    ' PushButton 430, 22, 100, 13, "Open DHS 2625", open_cs_2625_doc
				' PushButton 430, 122, 100, 13, "Open DHS 2707", open_cs_2707_doc
				' PushButton 430, 222, 100, 13, "Open DHS 7635", open_cs_7635_doc

			  Text 10, 370, 210, 10, "Confirm you have reviewed Hennepin County Information Information:"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If confirm_recap_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

			If ButtonPressed = verif_button Then
				Call verification_dialog
				err_msg = "LOOP"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

			' If ButtonPressed = open_cs_2625_doc OR ButtonPressed = open_cs_2707_doc OR ButtonPressed = open_cs_7635_doc Then
			' 	err_msg = "LOOP"
			' 	If ButtonPressed = open_cs_2625_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2625-ENG"
			' 	If ButtonPressed = open_cs_2707_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2707-ENG"
			' 	If ButtonPressed = open_cs_7635_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-7635-ENG"
			' End If
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
End If
save_your_work

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

' 'TESTING CODE - this is inplace so that the script doesn't error trying to update PROG.
' If update_revw = True OR update_prog = True Then
'
' 	Dialog1 = ""
' 	BeginDialog Dialog1, 0, 0, 246, 115, "Update Interview Date in STAT"
' 	  ButtonGroup ButtonPressed
' 	    OkButton 200, 95, 40, 15
' 	  Text 10, 10, 240, 10, "It appears STAT does not have the Interview Date coded into the panel."
' 	  Text 10, 20, 190, 10, "This makes sense, as you JUST completed the interview."
' 	  Text 10, 35, 215, 25, "We will be updating the script to do this for you, however, that functionality appears to be broken. So instead of making the script error all the time, we have removed the automatic functionality."
' 	  Text 10, 70, 225, 15, "You can update STAT now with the interview date or do it after the script run is complete, but it must be done manually for now."
' 	  Text 20, 90, 80, 10, "PROG Needs Update"
' 	  Text 20, 100, 80, 10, "REVW Needs Update"
' 	EndDialog
'
' 	dialog Dialog1
'
' End If
' update_revw = False
' update_prog = False

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

interview_time = ((timer - start_time) + add_to_time)/60
interview_time = Round(interview_time, 2)

intvw_done_msg_file = user_myDocs_folder & "interview done message.txt"
If user_ID_for_validation = "ERHO003" Then intvw_done_msg_file = user_c_drive_docs_folder & "interview done message.txt"
With (CreateObject("Scripting.FileSystemObject"))
	If .FileExists(intvw_done_msg_file) = True then .DeleteFile(intvw_done_msg_file)

	If .FileExists(intvw_done_msg_file) = False then
		Set objTextStream = .OpenTextFile(intvw_done_msg_file, 2, true)

		'Write the contents of the text file
		objTextStream.WriteLine "This interview has been COMPLETED!"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "The interview took " & interview_time & " minutes."
		objTextStream.WriteLine "The script is currently creating your PDF, SPEC/MEMO, and CASE/NOTEs. DO NOT TRY TO TAKE ANY ACTION ON THE COMPUTER WHILE THIS FINISHES."
		objTextStream.WriteLine "Agency Siganture is not required on the " & CAF_form & "."
		objTextStream.WriteLine ""
		objTextStream.WriteLine ""
		objTextStream.WriteLine "This is a great time to talk to the resident about: "
		objTextStream.WriteLine "  - The interview is complete."
		objTextStream.WriteLine "  - Advise of Next Steps."
		objTextStream.WriteLine "  - Ask if the resident has any final questions."
		objTextStream.WriteLine ""
		objTextStream.WriteLine "(This message will close once the script actions are finished.)"

		objTextStream.Close
	End If
End With
Set o2Exec = WshShell.Exec("notepad " & intvw_done_msg_file)


' complete_interview_msg = MsgBox("This interview is now completed and has taken " & interview_time & " minutes." & vbCr & vbCr & "The script will now create your interview notes in a PDF and enter CASE:NOTE(s) as needed.", vbInformation, "Interview Completed")

' script_end_procedure("At this point the script will create a PDF with all of the interview notes to save to ECF, enter a comprehensive CASE:NOTE, and update PROG or REVW with the interview date. Future enhancements will add more actions functionality.")
'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "CAF Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "CAF Form Details - CASE #" & MAXIS_case_number			'Title of the document
' objWord.Visible = True														'Let the worker see the document
objWord.Visible = False														'Let the worker see the document

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "NOTES on INTERVIEW"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

If MAXIS_case_number <> "" Then objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information
' If no_case_number_checkbox = checked Then objSelection.TypeText "New Case - no case number" & vbCr
objSelection.TypeText "Interview Date: " & interview_date & vbCR
objSelection.TypeText "DATE OF APPLICATION: " & CAF_datestamp & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR
objSelection.TypeText "Interview completed with: " & who_are_we_completing_the_interview_with & vbCR
objSelection.TypeText "Interview completed via: " & how_are_we_completing_the_interview & vbCR
length_of_interview = ((timer - start_time) + add_to_time)/60
length_of_interview = Round(length_of_interview, 2)
objSelection.TypeText "Interview length: " & length_of_interview & " minutes" & vbCR

If trim(interpreter_information) <> "" AND interpreter_information <> "No Interpreter Used" Then
	objSelection.TypeText "Interview had interpreter: " & interpreter_information & vbCr
	objSelection.TypeText "    Language: " & interpreter_language & vbCr
End If
If trim(arep_interview_id_information) <> "" Then objSelection.TypeText "AREP Identity Verification: " & arep_interview_id_information & vbCr
If trim(non_applicant_interview_info) <> "" Then objSelection.TypeText "Interviewee Information: " & non_applicant_interview_info & vbCr

objSelection.TypeText "Case Status at the time of interview: " & vbCR
If case_active = True Then objSelection.TypeText "   Case is ACTIVE" & vbCR
If case_pending = True Then objSelection.TypeText "   Case is PENDING" & vbCR
If case_pending = False AND case_active = False Then objSelection.TypeText "   Case is INACTIVE" & vbCR

Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 7, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objProgStatusTable = objDoc.Tables(1)		'Creates the table with the specific index'

objProgStatusTable.AutoFormat(16)							'This adds the borders to the table and formats it

objProgStatusTable.Cell(1, 1).SetHeight 15, 2
for row = 2 to 7
	objProgStatusTable.Cell(row, 1).SetHeight 12, 2			'setting the heights of the rows
Next

for row = 1 to 7
	objProgStatusTable.Rows(row).Cells.Split 1, 2, TRUE
Next

objProgStatusTable.Columns(1).Width = 150					'This sets the width of the table.
objProgStatusTable.Columns(2).Width = 200					'This sets the width of the table.
' objProgStatusTable.Columns(3).Width = 150					'This sets the width of the table.

'Now going to each cell and setting teh font size
objProgStatusTable.Cell(1, 1).Range.Font.Size = 11
objProgStatusTable.Cell(1, 2).Range.Font.Size = 11
For row = 2 to 7
	objProgStatusTable.Cell(row, 1).Range.Font.Size = 9
	objProgStatusTable.Cell(row, 2).Range.Font.Size = 9
Next

' objProgStatusTable.Cell(row, col).Range.Text =

objProgStatusTable.Cell(1, 1).Range.Text = "Program"
objProgStatusTable.Cell(1, 2).Range.Text = "Status"
' objProgStatusTable.Cell(1, 3).Range.Text = "Detail"

objProgStatusTable.Cell(2, 1).Range.Text = "SNAP"
objProgStatusTable.Cell(2, 2).Range.Text = snap_status
' objProgStatusTable.Cell(2, 3).Range.Text =
cash_col = 3
' If
If mfip_status <> "INACTIVE" Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "MFIP"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = mfip_status
	' objProgStatusTable.Cell(cash_col, 3).Range.Text = "MFIP"
	cash_col = cash_col + 1
End If
If dwp_status <> "INACTIVE" Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "DWP"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = dwp_status
	' objProgStatusTable.Cell(cash_col, 3).Range.Text = "MFIP"
	cash_col = cash_col + 1
End If
If ga_status <> "INACTIVE" Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "GA"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = ga_status
	' objProgStatusTable.Cell(cash_col, 3).Range.Text = "MFIP"
	cash_col = cash_col + 1
End If
If msa_status <> "INACTIVE" Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "MSA"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = msa_status
	' objProgStatusTable.Cell(cash_col, 3).Range.Text = "MFIP"
	cash_col = cash_col + 1
End If
If unknown_cash_pending = True Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "CASH"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = "PENDING"
	' objProgStatusTable.Cell(cash_col, 3).Range.Text = "CASH"
	cash_col = cash_col + 1
End If

If cash_col = 3 Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "CASH 1"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = "INACTIVE"
	cash_col = cash_col + 1
End If
If cash_col = 4 Then
	objProgStatusTable.Cell(cash_col, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = "INACTIVE"
	cash_col = cash_col + 1
End If

objProgStatusTable.Cell(5, 1).Range.Text = "GRH"
objProgStatusTable.Cell(5, 2).Range.Text = grh_status

objProgStatusTable.Cell(6, 1).Range.Text = "MA"
objProgStatusTable.Cell(6, 2).Range.Text = ma_status

objProgStatusTable.Cell(7, 1).Range.Text = "MSP"
objProgStatusTable.Cell(7, 2).Range.Text = msp_status

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeText vbCr
'Program CAF Information
caf_progs = ""
for the_memb = 0 to UBOUND(HH_MEMB_ARRAY, 2)
    If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
        If HH_MEMB_ARRAY(snap_req_checkbox, the_memb) = checked AND InStr(caf_progs, "SNAP") = 0 Then caf_progs = caf_progs & ", SNAP"
    	If HH_MEMB_ARRAY(cash_req_checkbox, the_memb) = checked AND InStr(caf_progs, "Cash") = 0 Then caf_progs = caf_progs & ", Cash"
    	If HH_MEMB_ARRAY(emer_req_checkbox, the_memb) = checked AND InStr(caf_progs, "EMER") = 0 Then caf_progs = caf_progs & ", EMER"
    End If
Next
If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
objSelection.TypeText "PROGRAMS REQUESTED ON CAF: " & caf_progs & vbCr
objSelection.Font.Size = "11"


'Ennumeration for SetHeight and SetWidth
'wdAdjustFirstColumn	2	Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
	' wdAdjustNone			0	Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
	' wdAdjustProportional	1	Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
	' wdAdjustSameWidth		3	Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.


objSelection.TypeText "PERSON 1 Information - Confirmed in the Interview"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 16, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objPers1Table = objDoc.Tables(2)		'Creates the table with the specific index'
'This table will be formatted to look similar to the structure of CAF Page 1

objPers1Table.AutoFormat(16)							'This adds the borders to the table and formats it
objPers1Table.Columns(1).Width = 500					'This sets the width of the table.

for row = 1 to 15 Step 2
	objPers1Table.Cell(row, 1).SetHeight 10, 2			'setting the heights of the rows
Next
for row = 2 to 16 Step 2
	objPers1Table.Cell(row, 1).SetHeight 15, 2
Next

'Now we are going to look at the the first and second rows. These have 4 cells to add details in and we will split the row into those 4 then resize them
For row = 1 to 2
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 140, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 85, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
'Now going to each cell and setting teh font size
For col = 1 to 4
	objPers1Table.Cell(1, col).Range.Font.Size = 6
	objPers1Table.Cell(2, col).Range.Font.Size = 12
Next

'Adding the headers
objPers1Table.Cell(1, 1).Range.Text = "APPLICANT'S LEGAL NAME - LAST"
objPers1Table.Cell(1, 2).Range.Text = "FIRST NAME"
objPers1Table.Cell(1, 3).Range.Text = "MIDDLE NAME"
objPers1Table.Cell(1, 4).Range.Text = "OTHER NAMES YOU USE"

'Adding the detail from the dialog
objPers1Table.Cell(2, 1).Range.Text = HH_MEMB_ARRAY(last_name_const, 0)
objPers1Table.Cell(2, 2).Range.Text = HH_MEMB_ARRAY(first_name_const, 0)
objPers1Table.Cell(2, 3).Range.Text = HH_MEMB_ARRAY(mid_initial, 0)
objPers1Table.Cell(2, 4).Range.Text = HH_MEMB_ARRAY(other_names, 0)

' objPers1Table.Cell(1, 2).Borders(wdBorderBottom).LineStyle = wdLineStyleNone			'commented out code dealing with borders
' objPers1Table.Cell(1, 3).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 4).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
' objPers1Table.Cell(1, 1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Range.Borders(9).LineStyle = 0
' objPers1Table.Rows(1).Borders(wdBorderBottom) = wdLineStyleNone

'Now formatting rows 3 and 4 - 3 is the header and 4 is the actual information
For row = 3 to 4
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 110, 2
	objPers1Table.Cell(row, 2).SetWidth 85, 2
	objPers1Table.Cell(row, 3).SetWidth 115, 2
	objPers1Table.Cell(row, 4).SetWidth 190, 2
Next
For col = 1 to 4
	objPers1Table.Cell(3, col).Range.Font.Size = 6
	objPers1Table.Cell(4, col).Range.Font.Size = 12
Next
'Adding the words to rows 3 and 4
objPers1Table.Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
objPers1Table.Cell(3, 2).Range.Text = "DATE OF BIRTH"
objPers1Table.Cell(3, 3).Range.Text = "GENDER"
objPers1Table.Cell(3, 4).Range.Text = "MARITAL STATUS"

objPers1Table.Cell(4, 1).Range.Text = HH_MEMB_ARRAY(ssn, 0)
objPers1Table.Cell(4, 2).Range.Text = HH_MEMB_ARRAY(date_of_birth, 0)
objPers1Table.Cell(4, 3).Range.Text = HH_MEMB_ARRAY(gender, 0)
objPers1Table.Cell(4, 4).Range.Text = HH_MEMB_ARRAY(marital_status, 0)

'Now formatting rows 5 and 6 - 5 is the header and 6 is the actual information
For row = 5 to 6
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(5, col).Range.Font.Size = 6
	objPers1Table.Cell(6, col).Range.Font.Size = 12
Next
'Adding the words to rows 5 and 6
objPers1Table.Cell(5, 1).Range.Text = "RESIDENCE ADDRESS - Confirmed in the Interview"
' objPers1Table.Cell(5, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(5, 2).Range.Text = "CITY"
objPers1Table.Cell(5, 3).Range.Text = "STATE"
objPers1Table.Cell(5, 4).Range.Text = "ZIP CODE"

If homeless_yn = "Yes" Then
	objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full & " - HOMELESS - "
Else
	objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full
End If
' objPers1Table.Cell(6, 2).Range.Text = ""
objPers1Table.Cell(6, 2).Range.Text = resi_addr_city
objPers1Table.Cell(6, 3).Range.Text = LEFT(resi_addr_state, 2)
objPers1Table.Cell(6, 4).Range.Text = resi_addr_zip

'Now formatting rows 7 and 8 - 7 is the header and 8 is the actual information
For row = 7 to 8
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 285, 2
	' objPers1Table.Cell(row, 2).SetWidth 55, 2
	objPers1Table.Cell(row, 2).SetWidth 110, 2
	objPers1Table.Cell(row, 3).SetWidth 35, 2
	objPers1Table.Cell(row, 4).SetWidth 70, 2
Next
For col = 1 to 4
	objPers1Table.Cell(7, col).Range.Font.Size = 6
	objPers1Table.Cell(8, col).Range.Font.Size = 12
Next
'Adding the words to rows 7 and 8
objPers1Table.Cell(7, 1).Range.Text = "MAILING ADDRESS"
' objPers1Table.Cell(7, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(7, 2).Range.Text = "CITY"
objPers1Table.Cell(7, 3).Range.Text = "STATE"
objPers1Table.Cell(7, 4).Range.Text = "ZIP CODE"

objPers1Table.Cell(8, 1).Range.Text = mail_addr_street_full
' objPers1Table.Cell(8, 2).Range.Text = ""
objPers1Table.Cell(8, 2).Range.Text = mail_addr_city
objPers1Table.Cell(8, 3).Range.Text = LEFT(mail_addr_state, 2)
objPers1Table.Cell(8, 4).Range.Text = mail_addr_zip

'Now formatting rows 9 and 10 - 9 is the header and 10 is the actual information
For row = 9 to 10
	objPers1Table.Rows(row).Cells.Split 1, 4, TRUE

	objPers1Table.Cell(row, 1).SetWidth 105, 2
	objPers1Table.Cell(row, 2).SetWidth 105, 2
	objPers1Table.Cell(row, 3).SetWidth 105, 2
	objPers1Table.Cell(row, 4).SetWidth 185, 2
Next
For col = 1 to 4
	objPers1Table.Cell(9, col).Range.Font.Size = 6
	objPers1Table.Cell(10, col).Range.Font.Size = 11
Next
'Adding the words to rows 9 and 10
objPers1Table.Cell(9, 1).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 2).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 3).Range.Text = "PHONE NUMBER"
objPers1Table.Cell(9, 4).Range.Text = "DO YOU LIVE ON A RESERVATION?"

'formatting the phone numbers so they all match and fit
Call format_phone_number(phone_one_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_two_number, "xxx-xxx-xxxx")
Call format_phone_number(phone_three_number, "xxx-xxx-xxxx")
If phone_one_type = "" OR phone_one_type = "Select One..." Then
	phone_one_info = phone_one_number
Else
	phone_one_info = phone_one_number & " (" & left(phone_one_type, 1) & ")"
End If

If phone_two_type = "" OR phone_two_type = "Select One..." Then
	phone_two_info = phone_two_number
Else
	phone_two_info = phone_two_number & " (" & left(phone_two_type, 1) & ")"
End If
If phone_three_type = "" OR phone_three_type = "Select One..." Then
	phone_three_info = phone_three_number
Else
	phone_three_info = phone_three_number & " (" & left(phone_three_type, 1) & ")"
End If
objPers1Table.Cell(10, 1).Range.Text = phone_one_info
objPers1Table.Cell(10, 2).Range.Text = phone_two_info
objPers1Table.Cell(10, 3).Range.Text = phone_three_info
objPers1Table.Cell(10, 4).Range.Text = reservation_yn & " - " & reservation_name

'Now formatting rows 11 and 12 - 11 is the header and 12 is the actual information
For row = 11 to 12
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 190, 2
	objPers1Table.Cell(row, 3).SetWidth 190, 2
Next
For col = 1 to 3
	objPers1Table.Cell(11, col).Range.Font.Size = 6
	objPers1Table.Cell(12, col).Range.Font.Size = 12
Next
'Adding the words to rows 11 and 12
objPers1Table.Cell(11, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
objPers1Table.Cell(11, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
objPers1Table.Cell(11, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

objPers1Table.Cell(12, 1).Range.Text = HH_MEMB_ARRAY(interpreter, 0)
objPers1Table.Cell(12, 2).Range.Text = HH_MEMB_ARRAY(spoken_lang, 0)
objPers1Table.Cell(12, 3).Range.Text = HH_MEMB_ARRAY(written_lang, 0)

'Now formatting rows 13 and 14 - 13 is the header and 14 is the actual information
For row = 13 to 14
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 120, 2
	objPers1Table.Cell(row, 2).SetWidth 270, 2
	objPers1Table.Cell(row, 3).SetWidth 110, 2
Next
For col = 1 to 3
	objPers1Table.Cell(13, col).Range.Font.Size = 6
	objPers1Table.Cell(14, col).Range.Font.Size = 12
Next
'Adding the words to rows 13 and 14
objPers1Table.Cell(13, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
objPers1Table.Cell(13, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
objPers1Table.Cell(13, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

objPers1Table.Cell(14, 1).Range.Text = HH_MEMB_ARRAY(last_grade_completed, 0)
objPers1Table.Cell(14, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(mn_entry_date, 0) & "   From: " & HH_MEMB_ARRAY(former_state, 0)
objPers1Table.Cell(14, 3).Range.Text = HH_MEMB_ARRAY(citizen, 0)

'Now formatting rows 15 and 16 - 15 is the header and 16 is the actual information
For row = 15 to 16
	objPers1Table.Rows(row).Cells.Split 1, 3, TRUE

	objPers1Table.Cell(row, 1).SetWidth 275, 2
	objPers1Table.Cell(row, 2).SetWidth 95, 2
	objPers1Table.Cell(row, 3).SetWidth 130, 2
Next
For col = 1 to 3
	objPers1Table.Cell(15, col).Range.Font.Size = 6
	objPers1Table.Cell(16, col).Range.Font.Size = 12
Next
'Adding the words to rows 15 and 16
objPers1Table.Cell(15, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
objPers1Table.Cell(15, 2).Range.Text = "ETHNICITY"
objPers1Table.Cell(15, 3).Range.Text = "RACE"

'defining a string that lists the programs based on the checkboxes of programs from the dialog'
If HH_MEMB_ARRAY(none_req_checkbox, 0) = checked then progs_applying_for = "NONE"
If HH_MEMB_ARRAY(snap_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", SNAP"
If HH_MEMB_ARRAY(cash_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Cash"
If HH_MEMB_ARRAY(emer_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

'defining a string of the races that were selected from checkboxes in the dialog.
If HH_MEMB_ARRAY(race_a_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Asian"
If HH_MEMB_ARRAY(race_b_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Black"
If HH_MEMB_ARRAY(race_n_checkbox, 0) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
If HH_MEMB_ARRAY(race_p_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
If HH_MEMB_ARRAY(race_w_checkbox, 0) = checked then race_to_enter = race_to_enter & ", White"
If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

objPers1Table.Cell(16, 1).Range.Text = progs_applying_for
objPers1Table.Cell(16, 2).Range.Text = HH_MEMB_ARRAY(ethnicity_yn, 0)
objPers1Table.Cell(16, 3).Range.Text = race_to_enter

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.TypeText "Household Lives in " & resi_addr_county & " County" & vbCR
If disc_out_of_county = "RESOLVED" Then objSelection.TypeText "- Household reported living Out of Hennepin County - Case Needs Transfer - additional interview conversation: " & disc_out_of_county_confirmation & vbCr

objSelection.TypeText "LIVING SITUATION: " & living_situation & vbCR
objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, 0) & vbCR
If disc_homeless_no_mail_addr = "RESOLVED" Then objSelection.TypeText "- Household Experiencing Housing Insecurity - MAIL is Primary Communication of Agency Requests and Actions - additional interview conversation: " & disc_homeless_confirmation & vbCr
If disc_no_phone_number = "RESOLVED" Then objSelection.TypeText "- No Phone Number was Provided - additional interview conversation: " & disc_phone_confirmation & vbCr

' objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF 1 - EXPEDITED QUESTIONS from the CAF"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
set objEXPTable = objDoc.Tables(3)		'Creates the table with the specific index'

objEXPTable.AutoFormat(16)							'This adds the borders to the table and formats it
objEXPTable.Columns(1).Width = 375					'Setting the widths of the columns
objEXPTable.Columns(2).Width = 120
for col = 1 to 2
	for row = 1 to 8
		objEXPTable.Cell(row, col).Range.Font.Bold = TRUE	'Making the cell text bold.
	next
next

'Adding the Expedited text to the table for Expedited
objEXPTable.Cell(1, 1).Range.Text = "1. How much income (cash or checks) did or will your household get this month?"
objEXPTable.Cell(1, 2).Range.Text = exp_q_1_income_this_month

objEXPTable.Cell(2, 1).Range.Text = "2. How much does your household (including children) have cash, checking or savings?"
objEXPTable.Cell(2, 2).Range.Text = exp_q_2_assets_this_month

objEXPTable.Cell(3, 1).Range.Text = "3. How much does your household pay for rent/mortgage per month?"
objEXPTable.Cell(3, 2).Range.Text = exp_q_3_rent_this_month

objEXPTable.Cell(4, 1).Range.Text = "   What utilities do you pay?"
If caf_exp_pay_heat_checkbox = checked Then util_pay = util_pay & "Heat, "
If caf_exp_pay_ac_checkbox = checked Then util_pay = util_pay & "Air Conditioning, "
If caf_exp_pay_electricity_checkbox = checked Then util_pay = util_pay & "Electricity, "
If caf_exp_pay_phone_checkbox = checked Then util_pay = util_pay & "Phone, "
If caf_exp_pay_none_checkbox = checked Then util_pay = util_pay & "NONE"
If right(util_pay, 2) = ", " Then util_pay = left(util_pay, len(util_pay) - 2)
objEXPTable.Cell(4, 2).Range.Text = util_pay

objEXPTable.Cell(5, 1).Range.Text = "4. Is anyone in your household a migrant or seasonal farm worker?"
objEXPTable.Cell(5, 2).Range.Text = exp_migrant_seasonal_formworker_yn

objEXPTable.Cell(6, 1).Range.Text = "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
objEXPTable.Cell(6, 2).Range.Text = exp_received_previous_assistance_yn

objEXPTable.Rows(7).Cells.Split 1, 6, TRUE										'Splitting the cells to add more detail for the three questions here
objEXPTable.Cell(7, 1).Range.Text = "When?"
objEXPTable.Cell(7, 2).Range.Text = exp_previous_assistance_when
objEXPTable.Cell(7, 3).Range.Text = "Where?"
objEXPTable.Cell(7, 4).Range.Text = exp_previous_assistance_where
objEXPTable.Cell(7, 5).Range.Text = "What?"
objEXPTable.Cell(7, 6).Range.Text = exp_previous_assistance_what

objEXPTable.Cell(8, 1).Range.Text = "6. Is anyone in your household pregnant?"
If exp_pregnant_who <> "" Then
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn & ", " &  exp_pregnant_who
Else
	objEXPTable.Cell(8, 2).Range.Text = exp_pregnant_yn
End If
objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.Font.Bold = TRUE
objSelection.TypeText "EXPEDITED Interview Answers:" & vbCr
objSelection.Font.Bold = FALSE
If case_is_expedited = True Then
	objSelection.TypeText "Based on income information this case APPEARS ELIGIBLE FOR EXPEDITED SNAP." & vbCr
Else
	objSelection.TypeText "This case does not appear eligible for expedited SNAP based on the income information." & vbCr
End If

objSelection.TypeText chr(9) & "Income in the month of application: " & intv_app_month_income & vbCr
objSelection.TypeText chr(9) & "Assets in the month of application: " & intv_app_month_asset & vbCr
objSelection.TypeText chr(9) & "Expenses in the month of application: " & app_month_expenses & vbCr
objSelection.TypeText chr(9) & chr(9) & "Housing expense in the month of application: " & intv_app_month_housing_expense & vbCr
objSelection.TypeText chr(9) & chr(9) & "Utilities in the month of application: " & utilities_cost & vbCr
If case_is_expedited = True Then
	If id_verif_on_file = "No" OR snap_active_in_other_state = "Yes" OR last_snap_was_exp = "Yes" Then
		objSelection.TypeText chr(9) & "Expedited Approval must be delayed:" & vbCr
		objSelection.TypeText chr(9) & chr(9) & "Detail: " & expedited_delay_info & vbCr
		If id_verif_on_file = "No" Then 			objSelection.TypeText chr(9) & chr(9) & "" & vbCr
		If snap_active_in_other_state = "Yes" Then 	objSelection.TypeText chr(9) & chr(9) & "" & vbCr
		If last_snap_was_exp = "Yes" Then 			objSelection.TypeText chr(9) & chr(9) & "" & vbCr
	End If
End If

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.Font.Bold = TRUE
objSelection.TypeText "Interview Answers:" & vbCr
objSelection.Font.Bold = FALSE
objSelection.TypeText chr(9) & "Identity: " & HH_MEMB_ARRAY(id_verif, 0) & vbCr
objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, 0) & vbCr
objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, 0) & vbCr
objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, 0) & vbCr
objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, 0) & vbCr
If HH_MEMB_ARRAY(client_verification_details, 0) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, 0) & vbCr

'Now we have a dynamic number of tables
'each table has to be defined with its index so we need to have a variable to increment
table_count = 4			'table index variable
additional_person = False
If UBound(HH_MEMB_ARRAY, 2) <> 0 Then
    For each_member = 1 to UBound(HH_MEMB_ARRAY, 2)
        If HH_MEMB_ARRAY(ignore_person, each_member) = False Then additional_person = True
    Next
End If
If additional_person = True Then
    numb_of_tables = 0
    For each_member = 1 to UBound(HH_MEMB_ARRAY, 2)
        If HH_MEMB_ARRAY(ignore_person, each_member) = False Then numb_of_tables = numb_of_tables + 1
    Next
    ReDim TABLE_ARRAY(numb_of_tables)		'defining the table array for as many persons aas are in the household - each person gets their own table
	array_counters = 0		'the incrementer for the table array'

	For each_member = 1 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, each_member) = False Then
            objSelection.TypeText "PERSON " & each_member + 1
    		Set objRange = objSelection.Range										'range is needed to create tables
    		objDoc.Tables.Add objRange, 10, 1										'This sets the rows and columns needed row then column'
    		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)			'Creates the table with the specific index - using the vairable index
    		table_count = table_count + 1											'incrementing the table index'

    		'This table is now formatted to match how the CAF looks with person information.
    		'This formatting uses 'spliting' and resizing to make theym look like the CAF
    		TABLE_ARRAY(array_counters).AutoFormat(16)								'This adds the borders to the table and formats it
    		TABLE_ARRAY(array_counters).Columns(1).Width = 500

    		for row = 1 to 9 Step 2
    			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
    		Next
    		for row = 2 to 10 Step 2
    			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
    		Next

    		For row = 1 to 2
    			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

    			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 140, 2
    			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 85, 2
    			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 85, 2
    			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
    		Next
    		For col = 1 to 4
    			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
    			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
    		Next

    		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "LEGAL NAME - LAST"
    		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "FIRST NAME"
    		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "MIDDLE NAME"
    		TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "OTHER NAMES"

    		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = HH_MEMB_ARRAY(last_name_const, each_member)
    		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = HH_MEMB_ARRAY(first_name_const, each_member)
    		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = HH_MEMB_ARRAY(mid_initial, each_member)
    		TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = HH_MEMB_ARRAY(other_names, each_member)

    		For row = 3 to 4
    			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 5, TRUE

    			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 95, 2
    			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 80, 2
    			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 65, 2
    			TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 190, 2
    			TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 70, 2
    		Next
    		For col = 1 to 5
    			TABLE_ARRAY(array_counters).Cell(3, col).Range.Font.Size = 6
    			TABLE_ARRAY(array_counters).Cell(4, col).Range.Font.Size = 12
    		Next
    		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "SOCIAL SECURITY NUMBER"
    		TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "DATE OF BIRTH"
    		TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "GENDER"
    		TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "RELATIONSHIP TO YOU"
    		TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "MARITAL STATUS"

    		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = HH_MEMB_ARRAY(ssn, each_member)
    		TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = HH_MEMB_ARRAY(date_of_birth, each_member)
    		TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = HH_MEMB_ARRAY(gender, each_member)
    		TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = HH_MEMB_ARRAY(rel_to_applcnt, each_member)
    		TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = Left(HH_MEMB_ARRAY(marital_status, each_member), 1)

    		For row = 5 to 6
    			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

    			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
    			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 190, 2
    			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 190, 2
    		Next
    		For col = 1 to 3
    			TABLE_ARRAY(array_counters).Cell(5, col).Range.Font.Size = 6
    			TABLE_ARRAY(array_counters).Cell(6, col).Range.Font.Size = 12
    		Next
    		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "DO YOU NEED AN INTERPRETER?"
    		TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "WHAT IS YOU PREFERRED SPOKEN LANGUAGE?"
    		TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "WHAT IS YOUR PREFERRED WRITTEN LANGUAGE?"

    		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = HH_MEMB_ARRAY(interpreter, each_member)
    		TABLE_ARRAY(array_counters).Cell(6, 2).Range.Text = HH_MEMB_ARRAY(spoken_lang, each_member)
    		TABLE_ARRAY(array_counters).Cell(6, 3).Range.Text = HH_MEMB_ARRAY(written_lang, each_member)

    		For row = 7 to 8
    			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

    			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 120, 2
    			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 270, 2
    			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
    		Next
    		For col = 1 to 3
    			TABLE_ARRAY(array_counters).Cell(7, col).Range.Font.Size = 6
    			TABLE_ARRAY(array_counters).Cell(8, col).Range.Font.Size = 12
    		Next
    		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "LAST SCHOOL GRADE COMPLETED"
    		TABLE_ARRAY(array_counters).Cell(7, 2).Range.Text = "MOST RECENTLY MOVED TO MINNESOTA"
    		TABLE_ARRAY(array_counters).Cell(7, 3).Range.Text = "US CITIZEN OR US NATIONAL?"

    		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = HH_MEMB_ARRAY(last_grade_completed, each_member)
    		TABLE_ARRAY(array_counters).Cell(8, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(mn_entry_date, each_member) & "   From: " & HH_MEMB_ARRAY(former_state, each_member)
    		TABLE_ARRAY(array_counters).Cell(8, 3).Range.Text = HH_MEMB_ARRAY(citizen, each_member)

    		For row = 9 to 10
    			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

    			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 275, 2
    			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 95, 2
    			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 130, 2
    		Next
    		For col = 1 to 3
    			TABLE_ARRAY(array_counters).Cell(9, col).Range.Font.Size = 6
    			TABLE_ARRAY(array_counters).Cell(10, col).Range.Font.Size = 12
    		Next
    		TABLE_ARRAY(array_counters).Cell(9, 1).Range.Text = "WHAT PROGRAMS ARE YOU APPLYING FOR?"
    		TABLE_ARRAY(array_counters).Cell(9, 2).Range.Text = "ETHNICITY"
    		TABLE_ARRAY(array_counters).Cell(9, 3).Range.Text = "RACE"

    		progs_applying_for = ""
    		If HH_MEMB_ARRAY(none_req_checkbox, each_member) = checked then progs_applying_for = "NONE"
    		If HH_MEMB_ARRAY(snap_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", SNAP"
    		If HH_MEMB_ARRAY(cash_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Cash"
    		If HH_MEMB_ARRAY(emer_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
    		If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

    		race_to_enter = ""
    		If HH_MEMB_ARRAY(race_a_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Asian"
    		If HH_MEMB_ARRAY(race_b_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Black"
    		If HH_MEMB_ARRAY(race_n_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
    		If HH_MEMB_ARRAY(race_p_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
    		If HH_MEMB_ARRAY(race_w_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", White"
    		If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

    		TABLE_ARRAY(array_counters).Cell(10, 1).Range.Text = progs_applying_for
    		TABLE_ARRAY(array_counters).Cell(10, 2).Range.Text = HH_MEMB_ARRAY(ethnicity_yn, each_member)
    		TABLE_ARRAY(array_counters).Cell(10, 3).Range.Text = race_to_enter


    		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

    		objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, each_member) & vbCR
    		' objSelection.Font.Bold = TRUE
    		' objSelection.TypeText "AGENCY USE:" & vbCr
    		' objSelection.Font.Bold = FALSE
    		objSelection.TypeText chr(9) & "Identity: " & HH_MEMB_ARRAY(id_verif, each_member) & vbCr
    		objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, each_member) & vbCr
    		objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, each_member) & vbCr
    		objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, each_member) & vbCr
    		objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, each_member) & vbCr
    		If HH_MEMB_ARRAY(client_verification_details, each_member) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, each_member) & vbCr

    		array_counters = array_counters + 1
        End If
	Next
Else
	objSelection.TypeText "THERE ARE NO OTHER PEOPLE TO BE LISTED ON THIS APPLICATION" & vbCr
	ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
End If

'This is the rest of the verbiage from the CAF. It is not kept in tables - for the most part
objSelection.TypeText "Q 1. Does everyone in your household buy, fix or eat food with you?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_1_yn & vbCr
If question_1_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_1_notes & vbCr
If question_1_verif_yn <> "Mot Needed" AND question_1_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_1_verif_yn & vbCr
If question_1_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_1_verif_details & vbCr
If question_1_yn <> "" OR trim(question_1_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_1_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_1_interview_notes & vbCR

objSelection.TypeText "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_2_yn & vbCr
If question_2_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_2_notes & vbCr
If question_2_verif_yn <> "Mot Needed" AND question_2_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_2_verif_yn & vbCr
If question_2_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_2_verif_details & vbCr
If question_2_yn <> "" OR trim(question_2_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_2_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_2_interview_notes & vbCR

objSelection.TypeText "Q 3. Is anyone in the household attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_3_yn & vbCr
If question_3_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_3_notes & vbCr
If question_3_verif_yn <> "Mot Needed" AND question_3_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_3_verif_yn & vbCr
If question_3_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_3_verif_details & vbCr
If question_3_yn <> "" OR trim(question_3_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_3_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_3_interview_notes & vbCR

objSelection.TypeText "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_4_yn & vbCr
If question_4_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_4_notes & vbCr
If question_4_verif_yn <> "Mot Needed" AND question_4_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_4_verif_yn & vbCr
If question_4_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_4_verif_details & vbCr
If question_4_yn <> "" OR trim(question_4_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_4_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_4_interview_notes & vbCR

objSelection.TypeText "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_5_yn & vbCr
If question_5_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_5_notes & vbCr
If question_5_verif_yn <> "Mot Needed" AND question_5_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_5_verif_yn & vbCr
If question_5_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_5_verif_details & vbCr
If question_5_yn <> "" OR trim(question_5_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_5_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_5_interview_notes & vbCR

objSelection.TypeText "Q 6. Is anyone unable to work for reasons other than illness or disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_6_yn & vbCr
If question_6_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_6_notes & vbCr
If question_6_verif_yn <> "Mot Needed" AND question_6_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_6_verif_yn & vbCr
If question_6_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_6_verif_details & vbCr
If question_6_yn <> "" OR trim(question_6_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_6_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_6_interview_notes & vbCR

objSelection.TypeText "Q 7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_7_yn & vbCr
If question_7_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_7_notes & vbCr
If question_7_verif_yn <> "Mot Needed" AND question_7_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_7_verif_yn & vbCr
If question_7_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_7_verif_details & vbCr
If question_7_yn <> "" OR trim(question_7_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_7_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_7_interview_notes & vbCR

objSelection.TypeText "Q 8. Has anyone in the household had a job or been self-employed in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_8_yn & vbCr
objSelection.TypeText "Q 8.a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?" & vbCr
objSelection.TypeText chr(9) & question_8a_yn & vbCr
If question_8_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_8_notes & vbCr
If question_8_verif_yn <> "Mot Needed" AND question_8_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_8_verif_yn & vbCr
If question_8_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_8_verif_details & vbCr
If question_8_yn <> "" OR trim(question_8_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_8_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_8_interview_notes & vbCR

objSelection.TypeText "Q 9. Does anyone in the household have a job or expect to get income from a job this month or next month?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_9_yn & vbCr
job_added = FALSE
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
		job_added = TRUE

		all_the_tables = UBound(TABLE_ARRAY) + 1
		ReDim Preserve TABLE_ARRAY(all_the_tables)
		Set objRange = objSelection.Range					'range is needed to create tables
		objDoc.Tables.Add objRange, 8, 1					'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		table_count = table_count + 1

		TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 400

		for row = 1 to 7 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
		Next
		for row = 2 to 8 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
		Next

		For row = 1 to 2
			TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

			TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 200, 2
			TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 90, 2
			TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
		Next
		For col = 1 to 3
			TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
			TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
		Next
		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Font.Size = 6
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Font.Size = 12
		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Font.Size = 6
		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Font.Size = 12
		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Font.Size = 6
		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Font.Size = 12

		TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "EMPLOYEE NAME"
		TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "HOURLY WAGE"
		TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "GROSS MONTHLY EARNINGS"
		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = JOBS_ARRAY(jobs_employee_name, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = JOBS_ARRAY(jobs_hourly_wage, each_job)
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)

		TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "EMPLOYER/BUSINESS NAME"
		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = JOBS_ARRAY(jobs_employer_name, each_job)

		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "CAF NOTES"
		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = JOBS_ARRAY(jobs_notes, each_job)

		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "INTERVIEW NOTES"
		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = JOBS_ARRAY(jobs_intv_notes, each_job)

		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		' objSelection.TypeParagraph()						'adds a line between the table and the next information

		array_counters = array_counters + 1

		objSelection.TypeText "Verification: " & JOBS_ARRAY(verif_yn, each_job) & " - " & JOBS_ARRAY(verif_details, each_job) & vbCR
	End If
next

If job_added = FALSE Then objSelection.TypeText chr(9) & "THERE ARE NO JOBS ENTERED." & vbCr

If question_9_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_9_notes & vbCr
' If question_9_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_9_verif_yn & vbCr
' If question_9_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_9_verif_details & vbCr

objSelection.TypeText "Q 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_10_yn & vbCr
If question_10_monthly_earnings <> "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: " & question_10_monthly_earnings & vbCr
If question_10_monthly_earnings = "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: NONE LISTED" & vbCr
If question_10_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_10_notes & vbCr
If question_10_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_10_verif_yn & vbCr
If question_10_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_10_verif_details & vbCr
If question_10_yn <> "" OR trim(question_10_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_10_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_10_interview_notes & vbCR

objSelection.TypeText "Q 11. Do you expect any changes in income, expenses or work hours?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_11_yn & vbCr
If question_11_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_11_notes & vbCr
If question_11_verif_yn <> "Mot Needed" AND question_11_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_11_verif_yn & vbCr
If question_11_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_11_verif_details & vbCr
If question_11_yn <> "" OR trim(question_11_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_11_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_11_interview_notes & vbCR

objSelection.Font.Bold = TRUE
objSelection.TypeText "Principal Wage Earner (PWE)" & vbCr
objSelection.Font.Bold = FALSE

all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 2					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
TABLE_ARRAY(array_counters).Columns(1).Width = 200
TABLE_ARRAY(array_counters).Columns(2).Width = 200

TABLE_ARRAY(array_counters).Cell(1, 1).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(1, 2).SetHeight 10, 2
TABLE_ARRAY(array_counters).Cell(2, 1).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(2, 2).SetHeight 15, 2
TABLE_ARRAY(array_counters).Cell(1, 1).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Font.Size = 6
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Font.Size = 12
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Font.Size = 12

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text ="DESIGNATED PWE"
TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text =pwe_selection
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text ="SIGNATURE OF APPLICANT"
' TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text ="VERBAL SIGNATURE"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

objSelection.TypeText "Q 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 5, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 500

For row = 1 to 4
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 75, 2
Next
TABLE_ARRAY(array_counters).Rows(5).Cells.Split 1, 3, TRUE

TABLE_ARRAY(array_counters).Cell(5, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(5, 2).SetWidth 175, 2
TABLE_ARRAY(array_counters).Cell(5, 3).SetWidth 75, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_12_rsdi_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "RSDI"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "$ " & question_12_rsdi_amt
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = question_12_ssi_yn
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = "SSI"
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "$ " & question_12_ssi_amt

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_12_va_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Veteran Benefits (VA)"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = "$ " & question_12_va_amt
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = question_12_ui_yn
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = "Unemployment Insurance"
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "$ " & question_12_ui_amt

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_12_wc_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Workers' Compensation"
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "$ " & question_12_wc_amt
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = question_12_ret_yn
TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = "Retirement Benefits"
TABLE_ARRAY(array_counters).Cell(3, 6).Range.Text = "$ " & question_12_ret_amt

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_12_trib_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Tribal payments"
TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = "$ " & question_12_trib_amt
TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = question_12_cs_yn
TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = "Child or Spousal support"
TABLE_ARRAY(array_counters).Cell(4, 6).Range.Text = "$ " & question_12_cs_amt

TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = question_12_other_yn
TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = "Other unearned income"
TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "$ " & question_12_other_amt

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_12_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_12_notes & vbCr
If question_12_verif_yn <> "Mot Needed" AND question_12_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_12_verif_yn & vbCr
If question_12_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_12_verif_details & vbCr
q_12_answered = FALSE
If question_12_rsdi_yn <> "" Then q_12_answered = TRUE
If question_12_rsdi_amt <> "" Then q_12_answered = TRUE
If question_12_ssi_yn <> "" Then q_12_answered = TRUE
If question_12_ssi_amt <> "" Then q_12_answered = TRUE
If question_12_va_yn <> "" Then q_12_answered = TRUE
If question_12_va_amt <> "" Then q_12_answered = TRUE
If question_12_ui_yn <> "" Then q_12_answered = TRUE
If question_12_ui_amt <> "" Then q_12_answered = TRUE
If question_12_wc_yn <> "" Then q_12_answered = TRUE
If question_12_wc_amt <> "" Then q_12_answered = TRUE
If question_12_ret_yn <> "" Then q_12_answered = TRUE
If question_12_ret_amt <> "" Then q_12_answered = TRUE
If question_12_trib_yn <> "" Then q_12_answered = TRUE
If question_12_trib_amt <> "" Then q_12_answered = TRUE
If question_12_cs_yn <> "" Then q_12_answered = TRUE
If question_12_cs_amt <> "" Then q_12_answered = TRUE
If question_12_other_yn <> "" Then q_12_answered = TRUE
If question_12_other_amt <> "" Then q_12_answered = TRUE
If question_12_notes <> "" Then q_12_answered = TRUE
If q_12_answered = TRUE Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_12_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_12_interview_notes & vbCR

objSelection.TypeText "Q 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_13_yn & vbCr
If question_13_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_13_notes & vbCr
If question_13_verif_yn <> "Mot Needed" AND question_13_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_13_verif_yn & vbCr
If question_13_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_13_verif_details & vbCr
If question_13_yn <> "" OR trim(question_13_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_13_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_13_interview_notes & vbCR

objSelection.TypeText "Q 14. Does your household have the following housing expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 3
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next
TABLE_ARRAY(array_counters).Rows(4).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(4, 1).SetWidth 90, 2
TABLE_ARRAY(array_counters).Cell(4, 2).SetWidth 430, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_14_rent_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Rent (include mobile home lot rental)"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_14_subsidy_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Rent or Section 8 subsidy"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_14_mortgage_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Mortgage/contract for deed payment"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_14_association_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Association fees"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_14_insurance_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Homeowner's insurance (if not included in mortgage) "
TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = question_14_room_yn
TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "Room and/or board"

TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = question_14_taxes_yn
TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Real estate taxes (if not included in mortgage)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_14_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_14_notes & vbCr
If question_14_verif_yn <> "Mot Needed" AND question_14_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_14_verif_yn & vbCr
If question_14_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_14_verif_details & vbCr
q_14_answered = FALSE
If question_14_rent_yn <> "" Then q_14_answered = TRUE
If question_14_subsidy_yn <> "" Then q_14_answered = TRUE
If question_14_mortgage_yn <> "" Then q_14_answered = TRUE
If question_14_association_yn <> "" Then q_14_answered = TRUE
If question_14_insurance_yn <> "" Then q_14_answered = TRUE
If question_14_room_yn <> "" Then q_14_answered = TRUE
If question_14_taxes_yn <> "" Then q_14_answered = TRUE
If q_14_answered = TRUE  Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_14_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_14_interview_notes & vbCR

objSelection.TypeText "Q 15. Does your household have the following utility expenses any time during the year?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 3, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 525

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 100, 2
	TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 75, 2
	TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 100, 2
Next
TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 2, TRUE

TABLE_ARRAY(array_counters).Cell(3, 1).SetWidth 75, 2
TABLE_ARRAY(array_counters).Cell(3, 2).SetWidth 450, 2

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_15_heat_ac_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Heating/air conditioning"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_15_electricity_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Electricity"
TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = question_15_cooking_fuel_yn
TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "Cooking fuel"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_15_water_and_sewer_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Water and sewer"
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_15_garbage_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Garbage removal"
TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = question_15_phone_yn
TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "Phone/cell phone"

TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = question_15_liheap_yn
TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_15_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_15_notes & vbCr
If question_15_verif_yn <> "Mot Needed" AND question_15_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_15_verif_yn & vbCr
If question_15_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_15_verif_details & vbCr
q_15_answered = FALSE
If question_15_heat_ac_yn <> "" Then q_15_answered = TRUE
If question_15_electricity_yn <> "" Then q_15_answered = TRUE
If question_15_cooking_fuel_yn <> "" Then q_15_answered = TRUE
If question_15_water_and_sewer_yn <> "" Then q_15_answered = TRUE
If question_15_garbage_yn <> "" Then q_15_answered = TRUE
If question_15_phone_yn <> "" Then q_15_answered = TRUE
If question_15_liheap_yn <> "" Then q_15_answered = TRUE
' If trim(question_15_phone_details) <> "" AND question_15_phone_details <> "Select or Type" Then q_15_answered = TRUE
If q_15_answered = TRUE  Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_15_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_15_interview_notes & vbCR
If trim(question_15_phone_details) <> "" AND question_15_phone_details <> "Select or Type" Then objSelection.TypeText chr(9) & "Detail about phone: " & question_15_phone_details & vbCr

objSelection.TypeText "Q 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_16_yn & vbCr
If question_16_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_16_notes & vbCr
If question_16_verif_yn <> "Mot Needed" AND question_16_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_16_verif_yn & vbCr
If question_16_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_16_verif_details & vbCr
If question_16_yn <> "" OR trim(question_16_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_16_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_16_interview_notes & vbCR

objSelection.TypeText "Q 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_17_yn & vbCr
If question_17_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_17_notes & vbCr
If question_17_verif_yn <> "Mot Needed" AND question_17_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_17_verif_yn & vbCr
If question_17_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_17_verif_details & vbCr
If question_17_yn <> "" OR trim(question_17_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_17_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_17_interview_notes & vbCR

objSelection.TypeText "Q 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_18_yn & vbCr
If question_18_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_18_notes & vbCr
If question_18_verif_yn <> "Mot Needed" AND question_18_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_18_verif_yn & vbCr
If question_18_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_18_verif_details & vbCr
If question_18_yn <> "" OR trim(question_18_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_18_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_18_interview_notes & vbCR

objSelection.TypeText "Q 19. For SNAP only: Does anyone in the household have medical expenses? " & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_19_yn & vbCr
If question_19_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_19_notes & vbCr
If question_19_verif_yn <> "Mot Needed" AND question_19_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_19_verif_yn & vbCr
If question_19_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_19_verif_details & vbCr
If question_19_yn <> "" OR trim(question_19_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_19_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_19_interview_notes & vbCR

objSelection.TypeText "Q 20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. " & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_20_cash_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Cash"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_20_acct_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Bank accounts (savings, checking, debit card, etc.)"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_20_secu_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Stocks, bonds, annuities, 401K, etc."
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_20_cars_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Vehicles (cars, trucks, motorcycles, campers, trailers)"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_20_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_20_notes & vbCr
If question_20_verif_yn <> "Mot Needed" AND question_20_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_20_verif_yn & vbCr
If question_20_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_20_verif_details & vbCr
q_20_answered = FALSE
If question_20_cash_yn <> "" Then q_20_answered = TRUE
If question_20_acct_yn <> "" Then q_20_answered = TRUE
If question_20_secu_yn <> "" Then q_20_answered = TRUE
If question_20_cars_yn <> "" Then q_20_answered = TRUE
If q_20_answered = TRUE  Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_20_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_20_interview_notes & vbCR

objSelection.TypeText "Q 21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_21_yn & vbCr
If question_21_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_21_notes & vbCr
If question_21_verif_yn <> "Mot Needed" AND question_21_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_21_verif_yn & vbCr
If question_21_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_21_verif_details & vbCr
If question_21_yn <> "" OR trim(question_21_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_21_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_21_interview_notes & vbCR

objSelection.TypeText "Q 22. For recertifications only: Did anyone move in or out of your home in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_22_yn & vbCr
If question_22_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_22_notes & vbCr
If question_22_verif_yn <> "Mot Needed" AND question_22_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_22_verif_yn & vbCr
If question_22_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_22_verif_details & vbCr
If question_22_yn <> "" OR trim(question_22_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_22_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_22_interview_notes & vbCR

objSelection.TypeText "Q 23. For children under the age of 19, are both parents living in the home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_23_yn & vbCr
If question_23_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_23_notes & vbCr
If question_23_verif_yn <> "Mot Needed" AND question_23_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_23_verif_yn & vbCr
If question_23_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_23_verif_details & vbCr
If question_23_yn <> "" OR trim(question_23_notes) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_23_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_23_interview_notes & vbCR

objSelection.TypeText "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?" & vbCr
all_the_tables = UBound(TABLE_ARRAY) + 1
ReDim Preserve TABLE_ARRAY(all_the_tables)
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
table_count = table_count + 1

'note that this table does not use an autoformat - which is why there are no borders on this table.'
TABLE_ARRAY(array_counters).Columns(1).Width = 520

For row = 1 to 2
	TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

	TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
	TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
	TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
Next

TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = question_24_rep_payee_yn
TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Representative Payee fees"
TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = question_24_guardian_fees_yn
TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Guardian or Conservator fees"

TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = question_24_special_diet_yn
TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Physician-prescribed special diet "
TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = question_24_high_housing_yn
TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "High housing costs"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
' objSelection.TypeParagraph()						'adds a line between the table and the next information

array_counters = array_counters + 1

If question_24_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_24_notes & vbCr
If question_24_verif_yn <> "Mot Needed" AND question_24_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_24_verif_yn & vbCr
If question_24_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_24_verif_details & vbCr
q_24_answered = FALSE
If question_24_rep_payee_yn <> "" Then q_24_answered = TRUE
If question_24_guardian_fees_yn <> "" Then q_24_answered = TRUE
If question_24_special_diet_yn <> "" Then q_24_answered = TRUE
If question_24_high_housing_yn <> "" Then q_24_answered = TRUE
If q_24_answered = TRUE  Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
If question_24_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_24_interview_notes & vbCR

objSelection.TypeText "CAF QUALIFYING QUESTIONS" & vbCr

objSelection.TypeText "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?" & vbCr
objSelection.TypeText chr(9) & qual_question_one & vbCr
If trim(qual_memb_one) <> "" AND qual_memb_one <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_one & vbCr
objSelection.TypeText "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?" & vbCr
objSelection.TypeText chr(9) & qual_question_two & vbCr
If trim(qual_memb_two) <> "" AND qual_memb_two <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_two & vbCr
objSelection.TypeText "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?" & vbCr
objSelection.TypeText chr(9) & qual_question_three & vbCr
If trim(qual_memb_there) <> "" AND qual_memb_there <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_there & vbCr
objSelection.TypeText "Has anyone in your household been convicted of a drug felony in the past 10 years?" & vbCr
objSelection.TypeText chr(9) & qual_question_four & vbCr
If trim(qual_memb_four) <> "" AND qual_memb_four <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_four & vbCr
objSelection.TypeText "Is anyone in your household currently violating a condition of parole, probation or supervised release?" & vbCr
objSelection.TypeText chr(9) & qual_question_five & vbCr
If trim(qual_memb_five) <> "" AND qual_memb_five <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_five & vbCr

objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

objSelection.TypeText "Signatures:" & vbCr
objSelection.Font.Size = "12"

objSelection.TypeText "Signature of Primary Adult: " & signature_detail
If signature_detail <> "Not Required" AND signature_detail <> "Blank" Then
	objSelection.TypeText " by " & signature_person & " on " & signature_date
End If
objSelection.TypeText vbCr

objSelection.TypeText "Signature of Secondary Adult: " & second_signature_detail
If second_signature_detail <> "Not Required" AND second_signature_detail <> "Blank" Then
	objSelection.TypeText " by " & second_signature_person & " on " & second_signature_date
End If
objSelection.TypeText vbCr
objSelection.TypeText vbCr

objSelection.Font.Size = "14"

objSelection.TypeText "AREP (Authorized Representative)" & vbCr
objSelection.Font.Size = "12"

If arep_action = "Yes - keep this AREP" Then
	If arep_in_MAXIS = True AND MAXIS_arep_updated = True Then
		objSelection.TypeText "AREP information in MAXIS changed/updated to:" & vbCR
		objSelection.TypeText "Name: " & arep_name & vbCr
		If trim(arep_relationship) <> "" AND arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & arep_relationship & vbCr
		If trim(arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & arep_phone_number & vbCr
		objSelection.TypeText "Address: " & arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip & vbCr
		If arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
		If arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
		If arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr

	ElseIf arep_in_MAXIS = False AND trim(arep_name) <> "" AND arep_on_CAF_checkbox = unchecked Then
		objSelection.TypeText "AREP information provided Verbally:" & vbCR
		objSelection.TypeText "Name: " & arep_name & vbCr
		If trim(arep_relationship) <> "" AND arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & arep_relationship & vbCr
		If trim(arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & arep_phone_number & vbCr
		objSelection.TypeText "Address: " & arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip & vbCr
		If arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
		If arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
		If arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr
	Else
		objSelection.TypeText "AREP Detail:" & vbCR
		objSelection.TypeText "Name: " & arep_name & vbCr
		If trim(arep_relationship) <> "" AND arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & arep_relationship & vbCr
		If trim(arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & arep_phone_number & vbCr
		objSelection.TypeText "Address: " & arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip & vbCr
		If arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
		If arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
		If arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr

	End If

	If arep_on_CAF_checkbox = checked Then objSelection.TypeText "This AREP information was entered on the CAF." & vbCR
ElseIf arep_action = "No - remove this AREP from my case" OR arep_authorization = "DO NOT AUTHORIZE AN AREP" Then
	objSelection.TypeText "AREP information known/provided but resident does NOT want this AREP to be Authorized:" & vbCR
	objSelection.TypeText "Name: " & arep_name & vbCr
	If trim(arep_relationship) <> "" AND arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & arep_relationship & vbCr
	If trim(arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & arep_phone_number & vbCr
	objSelection.TypeText "Address: " & arep_addr_street & " " & arep_addr_city & ", " & left(arep_addr_state, 2) & " " & arep_addr_zip & vbCr
	If arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
	If arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
	If arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr
End If
If arep_and_CAF_arep_match = False Then
	If trim(CAF_arep_name) <> "" AND CAF_arep_action = "Yes - add to MAXIS" Then
		objSelection.TypeText "AREP information provided on the CAF:" & vbCR
		objSelection.TypeText "Name: " & CAF_arep_name & vbCr
		If trim(CAF_arep_relationship) <> "" AND CAF_arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & CAF_arep_relationship & vbCr
		If trim(CAF_arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & CAF_arep_phone_number & vbCr
		objSelection.TypeText "Address: " & CAF_arep_addr_street & " " & CAF_arep_addr_city & ", " & left(CAF_arep_addr_state, 2) & " " & CAF_arep_addr_zip & vbCr
		If CAF_arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
		If CAF_arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
		If CAF_arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr
	ElseIf CAF_arep_action = "No - do not allow this AREP" OR arep_authorization = "DO NOT AUTHORIZE AN AREP" Then
		objSelection.TypeText "AREP information provided on the CAF but resident does NOT want this AREP to be Authorized:" & vbCR
		objSelection.TypeText "Name: " & CAF_arep_name & vbCr
		If trim(CAF_arep_relationship) <> "" AND CAF_arep_relationship <> "Select or Type" Then objSelection.TypeText "Relationship: " & CAF_arep_relationship & vbCr
		If trim(CAF_arep_phone_number) <> "" Then objSelection.TypeText "Phone: " & CAF_arep_phone_number & vbCr
		objSelection.TypeText "Address: " & CAF_arep_addr_street & " " & CAF_arep_addr_city & ", " & left(CAF_arep_addr_state, 2) & " " & CAF_arep_addr_zip & vbCr
		If CAF_arep_complete_forms_checkbox = checked Then objSelection.TypeText "Allow AREP to complete forms." & vbCr
		If CAF_arep_get_notices_checkbox = checked Then objSelection.TypeText "Send Notices and Mail to AREP." & vbCr
		If CAF_arep_use_SNAP_checkbox = checked Then objSelection.TypeText "Allow AREP to get and use SNAP Benefits." & vbCr
	End If
End If
If arep_authorization <> "Select One..." AND trim(arep_authorization) <> "" Then objSelection.TypeText "AREP AUTHORIZATION: " & arep_authorization & vbCr

If discrepancies_exist = True Then
	objSelection.TypeText vbCr
	objSelection.Font.Size = "14"

	objSelection.TypeText "Clarification on Possible Information Discrepancies" & vbCr
	objSelection.Font.Size = "12"

	If disc_no_phone_number = "RESOLVED" Then
		objSelection.TypeText "No Phone Number was Provided" & vbCr
		objSelection.TypeText "  - Resolution: " & disc_phone_confirmation & vbCr
	End If
	If disc_yes_phone_no_expense = "RESOLVED" Then
		objSelection.TypeText "Phone Number Listed but No Phone Expense" & vbCr
		objSelection.TypeText "  - Resolution: " & disc_yes_phone_no_expense_confirmation & vbCr
	End If
	If disc_no_phone_yes_expense = "RESOLVED" Then
		objSelection.TypeText "NO Phone Number Listed but Expense Exists" & vbCr
		objSelection.TypeText "  - Resolution: " & disc_no_phone_yes_expense_confirmation & vbCr
	End If
	If disc_homeless_no_mail_addr = "RESOLVED" Then
		objSelection.TypeText "Household Experiencing Housing Insecurity - MAIL is Primary Communication of Agency Requests and Actions" & vbCr
		objSelection.TypeText "  - Resolution: " & disc_homeless_confirmation & vbCr
	End If
	If disc_out_of_county = "RESOLVED" Then
		objSelection.TypeText "Household reported living Out of Hennepin County - Case Needs Transfer" & vbCr
		objSelection.TypeText "  - Resolution: " & disc_out_of_county_confirmation & vbCr
	End If
	If disc_rent_amounts = "RESOLVED" Then
		objSelection.TypeText "The Housing Expense information on CAF Page 1 and CAF Question 14 do not appear to Match" & vbCr
		objSelection.TypeText "  - CAF Page 1 Housing Expense: " & exp_q_3_rent_this_month & vbCr
		objSelection.TypeText "  - Question 14 Housing Expense: " & question_14_summary & vbCr
		objSelection.TypeText "  - Resolution: " & disc_rent_amounts_confirmation & vbCr
	End If
	If disc_utility_amounts = "RESOLVED" Then
		objSelection.TypeText "The Utility Expense information on CAF Page 1 and CAF Question 15 do not appear to Match" & vbCr
		objSelection.TypeText "  - CAF Page 1 Utility Expense: " & disc_utility_caf_1_summary & vbCr
		objSelection.TypeText "  - Question 15 Utility Expense: " & disc_utility_q_15_summary & vbCr
		objSelection.TypeText "  - Resolution: " & disc_utility_amounts_confirmation & vbCr
	End If
	objSelection.TypeText vbCr
End If
objSelection.TypeText vbCr
objSelection.Font.Size = "14"

objSelection.TypeText "--- VERIFICATIONS ---" & vbCr
objSelection.Font.Size = "12"

Call create_verifs_needed_list(verifs_selected, verifs_needed)
If trim(verifs_needed) <> "" Then
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
If trim(verifs_needed) = "" Then
	objSelection.TypeText "THERE ARE NO REQUESTED VERIFICATIONS INDICATED" & vbCr
Else
	objSelection.TypeText "Verifications Requested:" & vbCr
	If verif_req_form_sent_date <> "" Then objSelection.TypeText "Request sent on " & verif_req_form_sent_date & vbCr
	verif_counter = 1
	For each verif_item in verifs_array
		If number_verifs_checkbox = checked Then verif_item = verif_counter & ". " & verif_item
		objSelection.TypeText verif_item & vbCr
		verif_counter = verif_counter + 1
	Next
End If


'Here we are creating the file path and saving the file
file_safe_date = replace(date, "/", "-")		'dates cannot have / for a file name so we change it to a -

'We set the file path and name based on case number and date. We can add other criteria if important.
'This MUST have the 'pdf' file extension to work
pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"
If developer_mode = True Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Archive\TRAINING REGION Interviews - NOT for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"

'Now we save the document.
'MS Word allows us to save directly as a PDF instead of a DOC.
'the file path must be PDF
'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
objDoc.SaveAs pdf_doc_path, 17

'This looks to see if the PDF file has been correctly saved. If it has the file will exists in the pdf file path
If objFSO.FileExists(pdf_doc_path) = TRUE Then
	'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
	objDoc.Close wdDoNotSaveChanges
	objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)

	'Now we MEMO'
	Call start_a_new_spec_memo(memo_opened, False, "N", "N", "N", other_name, other_street, other_city, other_state, other_zip, False)

	CALL write_variable_in_SPEC_MEMO("You have completed your interview on " & interview_date)
	CALL write_variable_in_SPEC_MEMO("This is for the " & CAF_form_name & " you submiteed.")
	' Call write_variable_in_SPEC_MEMO("THERE ARE VERIFS")
	Call write_variable_in_SPEC_MEMO("In the interview, we reviewed a number of forms:")
	Call write_variable_in_SPEC_MEMO("  -Client Rights and Responsibilities (DHS 4163)")
	Call write_variable_in_SPEC_MEMO("  -How to use Your Minnesota EBT Card (DHS 3315A)")
	Call write_variable_in_SPEC_MEMO("  -Notice of Privacy Practices (DHS 3979)")
	Call write_variable_in_SPEC_MEMO("  -Notice About Income and Eligibility Verification System and Work Reporting System (DHS 2759)")
	Call write_variable_in_SPEC_MEMO("  -Appeal Rights and Civil Rights Notice (DHS 3353)")
	Call write_variable_in_SPEC_MEMO("  -Program Info for Cash, Food, and Child Care (DHS 2920)")
	Call write_variable_in_SPEC_MEMO("  -Domestic Violence Information (DHS 3477)")
	Call write_variable_in_SPEC_MEMO("  -Do you have a Disability? (DHS 4133)")
	' MFIP Cases Only (All on a single Dialog)
	If family_cash_case_yn = "Yes" Then
		Call write_variable_in_SPEC_MEMO("  -Reporting Responsibilities for MFIP (DHS 2647)")
		Call write_variable_in_SPEC_MEMO("  -Notice of Requirement to Attend MFIP Overview (DHS 2929)")
		Call write_variable_in_SPEC_MEMO("  -Family Violence Referral (DHS 3323)")
		' MFIP Cases with at least One Non-Custodial Parent (All on a single Dialog)
		If absent_parent_yn = "Yes" Then
			Call write_variable_in_SPEC_MEMO("  -Understanding Child Support-Handbook (DHS 3393)")
			Call write_variable_in_SPEC_MEMO("  -Referral to Support and Collections (DHS 3163B)")
			Call write_variable_in_SPEC_MEMO("  -Cooperation with Child Support Enforcement (DHS 2338)")
			If relative_caregiver_yn = "Y" Then Call write_variable_in_SPEC_MEMO("  -MFIP Child Only Assistance (DHS 5561)")
			' If Non-Custodial Caregiver -
		End If
		' MFIP Case with a Minor Caregiver (All on a single Dialog)
		If left(minor_caregiver_yn, 3) = "Yes" Then
			Call write_variable_in_SPEC_MEMO("  -Notice of Requirement to Attend School (DHS 2961)")
			' Call write_variable_in_SPEC_MEMO("  -Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS 2887)")
			Call write_variable_in_SPEC_MEMO("  -Graduate to Independence (DHS 2887)")
			If minor_caregiver_yn = "" Then Call write_variable_in_SPEC_MEMO("  -MFIP for Minor Caregivers (DHS 3238)")
		End If
	End If
	' SNAP Cases Only  (All on a single Dialog)
	If snap_case = True OR pend_snap_on_case = "Yes" Then
		Call write_variable_in_SPEC_MEMO("  -SNAP reporting responsibilities (DHS 2625)")
		Call write_variable_in_SPEC_MEMO("  -SNAP reporting responsibilities (DHS 2625)")
		' Call write_variable_in_SPEC_MEMO("  -Facts on Voluntarily Quitting Your Job If You Are on SNAP (DHS 2707)")
		Call write_variable_in_SPEC_MEMO("  -Facts on  Quitting Your Job When on SNAP (DHS 2707)")
		Call write_variable_in_SPEC_MEMO("  -Work Registration Notice (DHS 7635)")
	End If
	Call write_variable_in_SPEC_MEMO("The information on these forms can help you better understand and navigate public assistance programs. They are available online at mn.gov/dhs/general-public/publications-forms-resources/edocs or by calling us and requesting a copy mailed to you.")

	PF4

	If arep_authorization <> "DO NOT AUTHORIZE AN AREP" Then
		If arep_exists = True Then
			If arep_action = "Yes - keep this AREP" OR CAF_arep_action = "Yes - add to MAXIS" OR (arep_authorization <> "Select One..." AND arep_authorization <> "") Then
				Call start_a_new_spec_memo(memo_opened, False, "N", "N", "N", other_name, other_street, other_city, other_state, other_zip, False)
				CALL write_variable_in_SPEC_MEMO("You have indicated that you want to Authorize someone to be a Representative for your Public Assistance Case.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("An Authorized Representative is NOT a requirement for any Public Assistance, and you may remove this person from access to your case at any time.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("This is typically a person who can help you with gathering information, submitting documentation, and talking to the county and state on your behalf.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("You have authorized:")
				If CAF_arep_action = "Yes - add to MAXIS" Then
					CALL write_variable_in_SPEC_MEMO(CAF_arep_name)
				ElseIf arep_action = "Yes - keep this AREP" Then
					CALL write_variable_in_SPEC_MEMO(arep_name)
				ElseIf trim(CAF_arep_name) <> "" AND CAF_arep_action <> "No - do not allow this AREP" Then
					CALL write_variable_in_SPEC_MEMO(CAF_arep_name)
				ElseIf trim(arep_name) <> "" AND arep_action <> "No - remove this AREP from my case" Then
					CALL write_variable_in_SPEC_MEMO(arep_name)
				End If
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("If you have any questions, or to remove an AREP, call the county.")

				PF4

				Call start_a_blank_CASE_NOTE
				Call write_variable_in_CASE_NOTE("AREP REQUESTED ADDED TO CASE")
				Call write_variable_in_CASE_NOTE("In the Interview, Resident requested AREP added to the case.")
				If CAF_arep_action = "Yes - add to MAXIS" Then
					CALL write_bullet_and_variable_in_CASE_NOTE("Name", CAF_arep_name)
					CALL write_bullet_and_variable_in_CASE_NOTE("Relationship", CAF_arep_relationship)
					CALL write_bullet_and_variable_in_CASE_NOTE("Phone Number", CAF_arep_phone_number)
					CALL write_bullet_and_variable_in_CASE_NOTE("Street Address", CAF_arep_addr_street)
					CALL write_bullet_and_variable_in_CASE_NOTE("City", CAF_arep_addr_city)
					CALL write_bullet_and_variable_in_CASE_NOTE("State", CAF_arep_addr_state)
					CALL write_bullet_and_variable_in_CASE_NOTE("Zip", CAF_arep_addr_zip)

					If CAF_arep_complete_forms_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can fill out Forms.")
					If CAF_arep_get_notices_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP should get Notices.")
					If CAF_arep_use_SNAP_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can get and use SNAP Benefits.")
				ElseIf arep_action = "Yes - keep this AREP" Then
					CALL write_bullet_and_variable_in_CASE_NOTE("Name", arep_name)
					CALL write_bullet_and_variable_in_CASE_NOTE("Relationship", arep_relationship)
					CALL write_bullet_and_variable_in_CASE_NOTE("Phone Number", arep_phone_number)
					CALL write_bullet_and_variable_in_CASE_NOTE("Street Address", arep_addr_street)
					CALL write_bullet_and_variable_in_CASE_NOTE("City", arep_addr_city)
					CALL write_bullet_and_variable_in_CASE_NOTE("State", arep_addr_state)
					CALL write_bullet_and_variable_in_CASE_NOTE("Zip", arep_addr_zip)

					If arep_complete_forms_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can fill out Forms.")
					If arep_get_notices_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP should get Notices.")
					If arep_use_SNAP_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can get and use SNAP Benefits.")
				ElseIf trim(CAF_arep_name) <> "" AND CAF_arep_action <> "No - do not allow this AREP" Then
					CALL write_bullet_and_variable_in_CASE_NOTE("Name", CAF_arep_name)
					CALL write_bullet_and_variable_in_CASE_NOTE("Relationship", CAF_arep_relationship)
					CALL write_bullet_and_variable_in_CASE_NOTE("Phone Number", CAF_arep_phone_number)
					CALL write_bullet_and_variable_in_CASE_NOTE("Street Address", CAF_arep_addr_street)
					CALL write_bullet_and_variable_in_CASE_NOTE("City", CAF_arep_addr_city)
					CALL write_bullet_and_variable_in_CASE_NOTE("State", CAF_arep_addr_state)
					CALL write_bullet_and_variable_in_CASE_NOTE("Zip", CAF_arep_addr_zip)

					If CAF_arep_complete_forms_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can fill out Forms.")
					If CAF_arep_get_notices_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP should get Notices.")
					If CAF_arep_use_SNAP_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can get and use SNAP Benefits.")
				ElseIf trim(arep_name) <> "" AND arep_action <> "No - remove this AREP from my case" Then
					CALL write_bullet_and_variable_in_CASE_NOTE("Name", arep_name)
					CALL write_bullet_and_variable_in_CASE_NOTE("Relationship", arep_relationship)
					CALL write_bullet_and_variable_in_CASE_NOTE("Phone Number", arep_phone_number)
					CALL write_bullet_and_variable_in_CASE_NOTE("Street Address", arep_addr_street)
					CALL write_bullet_and_variable_in_CASE_NOTE("City", arep_addr_city)
					CALL write_bullet_and_variable_in_CASE_NOTE("State", arep_addr_state)
					CALL write_bullet_and_variable_in_CASE_NOTE("Zip", arep_addr_zip)

					If arep_complete_forms_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can fill out Forms.")
					If arep_get_notices_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP should get Notices.")
					If arep_use_SNAP_checkbox = checked Then Call write_variable_in_CASE_NOTE("* AREP can get and use SNAP Benefits.")
				End If
				If arep_authorization = "Select One..." Then arep_authorization = ""
				Call write_bullet_and_variable_in_CASE_NOTE("AREP AUTHORIZATION", arep_authorization)
				Call write_variable_in_CASE_NOTE("---")
				Call write_variable_in_CASE_NOTE(worker_signature)
			End If
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




		' Call start_a_blank_CASE_NOTE
		'
	    ' If IsDate(snap_denial_date) = TRUE Then
	    '     case_note_header_text = "Expedited Determination: SNAP to be denied"
	    ' Else
	    '     IF case_is_expedited = True then
	    '     	case_note_header_text = "Expedited Determination: SNAP appears expedited"
	    '     ELSEIF case_is_expedited = False then
	    '     	case_note_header_text = "Expedited Determination: SNAP does not appear expedited"
	    '     END IF
	    ' End If
	    ' Call write_variable_in_CASE_NOTE(case_note_header_text)
	    ' If interview_date <> "" Then Call write_variable_in_case_note ("* Interview completed on: " & interview_date & " and full Expedited Determination Done")
	    ' If IsDate(snap_denial_date) = TRUE Then
	    '     Call write_variable_in_CASE_NOTE("* SNAP to be denied on " & snap_denial_date & ". Since case is not SNAP eligible, case cannot receive Expedited issuance.")
	    '     If case_is_expedited = True Then
	    '         Call write_variable_with_indent_in_CASE_NOTE("Case is determined to meet criteria based upon income alone.")
	    '         Call write_variable_with_indent_in_CASE_NOTE("Expedited approval requires case to be otherwise eligble for SNAP and this does not meet this criteria.")
	    '     ElseIf case_is_expedited = False Then
	    '         Call write_variable_with_indent_in_CASE_NOTE("Expedited SNAP cannot be approved as case does not meet all criteria")
	    '     End If
	    '     Call write_bullet_and_variable_in_CASE_NOTE("Explanation of Denial", snap_denial_explain)
	    ' Else
	    '     IF case_is_expedited = True Then
	    '         If trim(exp_snap_approval_date) <> "" Then
	    '             Call write_variable_in_case_note ("* Case is determined to meet criteria and Expedited SNAP can be approved.")
	    '         Else
	    '             Call write_variable_in_case_note ("* Case is determined to meet expedited SNAP criteria, approval not yet completed.")
	    '         End If
	    '     End If
	    '     IF case_is_expedited = False Then Call write_variable_in_case_note ("* Expedited SNAP cannot be approved as case does not meet all criteria")
	    '     If case_is_expedited = True Then
	    '         If IsDate(exp_snap_approval_date) = TRUE Then Call write_variable_in_CASE_NOTE("* SNAP EXP approved on " & exp_snap_approval_date & " - " & DateDiff("d", CAF_datestamp, exp_snap_approval_date) & " days after the date of application.")
	    '         Call write_bullet_and_variable_in_CASE_NOTE("Reason for delay", exp_snap_delays)
	    '     End If
	    ' End If
	    ' If trim(intv_app_month_income) <> "" OR trim(intv_app_month_asset) <> "" OR trim(app_month_expenses) <> "" Then
	    '     Call write_variable_in_CASE_NOTE("* Expedited Determination is based on information from application month:")
	    '     Call write_variable_with_indent_in_CASE_NOTE("Income: $" & intv_app_month_income)
	    '     Call write_variable_with_indent_in_CASE_NOTE("Assets: $" & intv_app_month_asset)
	    '     Call write_variable_with_indent_in_CASE_NOTE("Expenses (Shelter & Utilities): $" & app_month_expenses)
	    ' End If
		'
	    ' Call write_variable_in_CASE_NOTE("---")
	    ' Call write_variable_in_CASE_NOTE(worker_signature)

	    PF3
	End If

	qual_questions_yes = FALSE
	If qual_question_one = "Yes" Then qual_questions_yes = TRUE
	If qual_question_two = "Yes" Then qual_questions_yes = TRUE
	If qual_question_three = "Yes" Then qual_questions_yes = TRUE
	If qual_question_four = "Yes" Then qual_questions_yes = TRUE
	If qual_question_five = "Yes" Then qual_questions_yes = TRUE

	If qual_questions_yes = TRUE Then
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

	'setting the end message
	end_msg = "Success! The information you have provided about the interview and all of the notes have been saved in a PDF. This PDF will be uploaded to ECF by the ES Support Team for Case # " & MAXIS_case_number & " and will remain in the CASE RECORD. CASE:NOTES have also been entered with the full interview detail."
	o2Exec.Terminate()
	'Now we ask if the worker would like the PDF to be opened by the script before the script closes
	'This is helpful because they may not be familiar with where these are saved and they could work from the PDF to process the reVw
	reopen_pdf_doc_msg = MsgBox("The information gathered in the interview has been saved as a PDF and will be added to ECF as a separate 'Interview Notes' document." & vbCr & vbCr & "This document will take the place of your CAF INTERVIEW ANNOTATIONS, as long as you have entered all interview notes to the script." & vbCr & "Agency Signature is not required on the application form." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbSystemModal + vbYesNo, "Open PDF Doc?")
	If reopen_pdf_doc_msg = vbYes Then
		With (CreateObject("Scripting.FileSystemObject"))

			If .FileExists(pdf_doc_path) = TRUE Then
				run_path = chr(34) & pdf_doc_path & chr(34)
				wshshell.Run run_path
				end_msg = end_msg & vbCr & vbCr & "The PDF has been opened for you to view the information that has been saved."
			Else
				end_msg = end_msg & vbCr & vbCr & "The script could not open the PDF document because the file could not be found." & vbCr & "This may be because the file is already being worked on by ES Support Team, or there could be a slight network connection slowdown. If you still need the PDF opened, you can try UTILITIES - Open Interview PDF to attempt to open the file, or check ECF to see if the document has already been added."
			End If
		End With
	End If
Else
	o2Exec.Terminate()
	end_msg = "Something has gone wrong - the CAF information has NOT been saved correctly to be processed." & vbCr & vbCr & "You can either save the Word Document that has opened as a PDF in the Assignment folder OR Close that document without saving and RERUN the script. Your details have been saved and the script can reopen them and attampt to create the files again. When the script is running, it is best to not interrupt the process."
End If

end_msg = end_msg & vbCr & vbCr & "The documment created for the ECF Case File can serve in place of any annotations as long as you entered all of your interview notes into the script. If you have entered all of the interview notes for this interview, there is no need to annotate the application form in ECF."
end_msg = end_msg & vbCr & vbCr & "Hennepin County does not require an Agency Signature to be added to the application form. Details can be found in the HSR Manual: https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Applications.aspx (Search: Applications)."
With (CreateObject("Scripting.FileSystemObject"))
	.DeleteFile(intvw_done_msg_file)
End With

revw_pending_table = False                                                      'Determining if we should be adding this case to the CasesPending SQL Table
If unknown_cash_pending = True Then revw_pending_table = True                   'case should be pending cash or snap and NOT have SNAP active
If ga_status = "PENDING" Then revw_pending_table = True
If msa_status = "PENDING" Then revw_pending_table = True
If mfip_status = "PENDING" Then revw_pending_table = True
If dwp_status = "PENDING" Then revw_pending_table = True
If grh_status = "PENDING" Then revw_pending_table = True
If snap_status = "PENDING" Then revw_pending_table = True
If snap_status = "ACTIVE" Then revw_pending_table = False

'Here we go to ensure this case is listed in the CasesPending table for ES Workflow
If developer_mode = False AND revw_pending_table = True Then                    'Only do this if not in training region.

    eight_digit_case_number = right("00000000"&MAXIS_case_number, 8)            'The SQL table functionality needs the leading 0s added to the Case Number

    If unknown_cash_pending = True Then cash_stat_code = "P"                    'determining the program codes for the table entry

    If ma_status = "INACTIVE" Or ma_status = "APP CLOSE" Then hc_stat_code = "I"
    If ma_status = "ACTIVE" Or ma_status = "APP OPEN" Then hc_stat_code = "A"
    If ma_status = "REIN" Then hc_stat_code = "R"
    If ma_status = "PENDING" Then hc_stat_code = "P"
    If msp_status = "INACTIVE" Or msp_status = "APP CLOSE" Then hc_stat_code = "I"
    If msp_status = "ACTIVE" Or msp_status = "APP OPEN" Then hc_stat_code = "A"
    If msp_status = "REIN" Then hc_stat_code = "R"
    If msp_status = "PENDING" Then hc_stat_code = "P"
    If unknown_hc_pending = True Then hc_stat_code = "P"

    If ga_status = "PENDING" Then ga_stat_code = "P"
    If ga_status = "REIN" Then ga_stat_code = "R"
    If ga_status = "ACTIVE" Or ga_status = "APP OPEN" Then ga_stat_code = "A"
    If ga_status = "INACTIVE" Or ga_status = "APP CLOSE" Then ga_stat_code = "I"

    If grh_status = "PENDING" Then grh_stat_code = "P"
    If grh_status = "REIN" Then grh_stat_code = "R"
    If grh_status = "ACTIVE" Or grh_status = "APP OPEN" Then grh_stat_code = "A"
    If grh_status = "INACTIVE" Or grh_status = "APP CLOSE" Then grh_stat_code = "I"

    If emer_status = "PENDING" Then emer_stat_code = "P"
    If emer_status = "REIN" Then emer_stat_code = "R"
    If emer_status = "ACTIVE" Or emer_status = "APP OPEN" Then emer_stat_code = "A"
    If emer_status = "INACTIVE" Or emer_status = "APP CLOSE" Then emer_stat_code = "I"

    If mfip_status = "PENDING" Then mfip_stat_code = "P"
    If mfip_status = "REIN" Then mfip_stat_code = "R"
    If mfip_status = "ACTIVE" Or mfip_status = "APP OPEN" Then mfip_stat_code = "A"
    If mfip_status = "INACTIVE" Or mfip_status = "APP CLOSE" Then mfip_stat_code = "I"

    If snap_status = "PENDING" Then snap_stat_code = "P"
    If snap_status = "REIN" Then snap_stat_code = "R"
    If snap_status = "ACTIVE" Or snap_status = "APP OPEN" Then snap_stat_code = "A"
    If snap_status = "INACTIVE" Or snap_status = "APP CLOSE" Then snap_stat_code = "I"

    appears_expedited_for_data_table = 1                                        'Setting if case is Expedited or not based on information in the Determination.
    If is_elig_XFS = False Then appears_expedited_for_data_table = 0

    If IsDate(CAF_datestamp) = True Then CAF_datestamp = DateAdd("d", 0, CAF_datestamp)     'make sure that CAF date is formatted as a date

    'Setting constants
    Const adOpenStatic = 3
    Const adLockOptimistic = 3

    'Creating objects for Access
    Set objConnection = CreateObject("ADODB.Connection")
    Set objRecordSet = CreateObject("ADODB.Recordset")

    'This is the BZST connection to SQL Database'
    objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""

    'delete a record if the case number matches
    objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection

    'if one was found we are going to delete that record
    If current_case_record_found = True Then objRecordSet.Open "DELETE FROM ES.ES_CasesPending WHERE CaseNumber = '" & eight_digit_case_number & "'", objConnection

    'Add a new record with this case information'
    objRecordSet.Open "INSERT INTO ES.ES_CasesPending (WorkerID, CaseNumber, CaseName, ApplDate, FSStatusCode, CashStatusCode, HCStatusCode, GAStatusCode, GRStatusCode, EAStatusCode, MFStatusCode, IsExpSnap, UpdateDate)" &  _
                      "VALUES ('" & worker_id_for_data_table & "', '" & eight_digit_case_number & "', '" & case_name_for_data_table & "', '" & CAF_datestamp & "', '" & snap_stat_code & "', '" & cash_stat_code & "', '" & hc_stat_code & "', '" & ga_stat_code & "', '" & grh_stat_code & "', '" & emer_stat_code & "', '" & mfip_stat_code & "', '" & appears_expedited_for_data_table & "', '" & date & "')", objConnection, adOpenStatic, adLockOptimistic

    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
End If

Call script_end_procedure_with_error_report(end_msg)


'POLICY NOTES
'
' Here is what Ann from Internal Services said about additional training:
'
' There is a training in IPAM that covers how to interview and covers annotating.
'
' Per CM
' WHAT IS A COMPLETE APPLICATION (state.mn.us)
' https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00051203
' obtain the answers from the client at the time of the interview and clearly document the information provided.
'
' APPLICATION INTERVIEWS (state.mn.us)
' https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00051212
' Nothing mentioned in this section either
'
' IPAM
' An Eligibility Workers Guide to the Combined Application Form (With Answers).pdf (state.mn.us)
' https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/_layouts/15/WopiFrame.aspx?sourcedoc=%7B3230AF4F-4FA7-448C-BAA7-506671E03A49%7D&file=An%20Eligibility%20Workers%20Guide%20to%20the%20Combined%20Application%20Form%20(With%20Answers).pdf&action=default&IsList=1&ListId=%7B032C9304-E9F4-4ED6-90A0-92F9CC18CD31%7D&ListItemId=2
' Answer section page 64
' 1) On what form do you record information from the interview?
' Information from the interview must be recorded on the CAF and in MAXIS CASE/NOTES, in sufficient detail for other workers and supervisors to follow the adequacy of the certification process and the accuracy of your decisions.
