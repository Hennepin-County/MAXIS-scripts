'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

'Handling for if run directly instead of the through the power pad.
If db_full_string = "" Then
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-Scripts\locally-installed-files\SETTINGS - GLOBAL VARIABLES.vbs")
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
End If

If script_repository = "" Then
	script_repository = "C:\MAXIS-Scripts\"		'This is a safety measure for if the script is run directly.
	run_locally = True
End If

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

Call select_testing_file("ALL", "", "notes/interview.vbs", "more-q4-interview-updates", False, False)


'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/05/2025", "Update the display of person information and adjustment to how person details needed for the interview are gathered.", "Casey Love, Hennepin County")
call changelog_update("06/23/2025", "The CAF form question were updated by DHS. The script has been updated to align with this new CAF layout and question format.##~## ##~##NOTE - ANY INTERVIEW DETAILS SAVED PRIOR TO TODAY WILL NOT BE ABLE TO BE RESTORED.##~## ##~##The interview details restoration has been updated to ensure the same form and version were selected for the information to be restored.", "Casey Love, Hennepin County")
call changelog_update("05/29/2025", "Condensed information displayed on EXPEDITED dialog to reduce risk of information extending past edges of dialog", "Mark Riegel, Hennepin County")
call changelog_update("05/02/2025", "Updated button locations for verifications dialog", "Mark Riegel, Hennepin County")
call changelog_update("01/27/2025", "Interview Updates:##~## - Added a 'Clear ALL' button to verifications.##~##   (New Interview - Verifications instruction document!)##~## - Remove entry of signature date as it is not necessary to document.##~## - Added information on the WIF and CASE/NOTE about verbal signature to align with policy.##~## - Updated some formatting and verbiage to align with different form types.##~## - Fixed bug in the Expedited information in the WIF.##~## ##~##These updates are far reaching and with a large script like the Interview script, there may be places where additional functionality or updates are needed. Please report anything you notice about these changes.##~##", "Casey Love, Hennepin County")
call changelog_update("11/27/2024", "Update to the process for documenting verbal program requests and documenting verbal program request withdrawals.##~##", "Casey Love, Hennepin County")
call changelog_update("11/20/2024", "BIG NEWS ! ##~## ##~## The Interview Script now supports different questions for different form selections!!!!##~## ##~##As this is brand new AND a very large change there may be some unexpected results or bugs. Please alert the script team to any bugs, questions, or thoughts you have on the updates.##~## ##~##We are very excited to be able to get this functionality out.", "Casey Love, Hennepin County")
call changelog_update("10/23/2024", "Added emergency questions.", "Mark Riegel, Hennepin County")
call changelog_update("09/30/2023", "Revised rights and responsibilities dialogs.", "Megan Geissler, Hennepin County")
call changelog_update("08/30/2023", "Updates to NOTES - Interview##~##- The script will carry over 'Interview Notes' from relevant questions into the Expedited Determination so you can reference notes already made during the script run.##~##- Additional resources around EBT Cards, including reference to the HSR Manual Accounting page.##~##- Background updates and bug fixes.##~## ##~##Please send any comments or feedback about these update to hsph.ews.bluezonescripts@hennepin.us.", "Casey Love, Hennepin County")
call changelog_update("07/21/2023", "Updated function that sends an email through Outlook", "Mark Riegel, Hennepin County")
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
const remo_info_const               = 69
const requires_update               = 70
const last_const					= 71

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

	If expedited_screening_on_form = False Then expedited_screening = ""
end function

full_err_msg = full_err_msg & "~!~" & "1^* CAF DATESTAMP ##~##   - Enter a valid date for the CAF datestamp.##~##"

function check_for_errors(interview_questions_clear)
	' If  Then err_msg = err_msg & "~!~" & "1^* FIELD##~##   - "
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

    page_3_errors = False
    all_members_in_MN_notes = trim(all_members_in_MN_notes)
    If all_members_in_MN_yn = "" Then err_msg = err_msg & "~!~" & "3 ^* Do ALL Members intend to reside in MN?##~##   - Indicate if all household members intend to reside in Minnesota."
    If all_members_in_MN_yn = "No" and all_members_in_MN_notes = "" Then err_msg = err_msg & "~!~" & "3 ^* Do ALL Members intend to reside in MN? - Notes##~##   - Since it is indicated that not all household members intend to reside in Minnesota, add notes to explain the details of member residence."
    anyone_pregnant_notes = trim(anyone_pregnant_notes)
    If anyone_pregnant_yn = "Yes" and anyone_pregnant_notes = "" Then err_msg = err_msg & "~!~" & "3 ^* Is anyone pregnant? - Notes##~##   - Since it is indicated that someone is pregnant, add notes with additional details."
    If anyone_pregnant_yn = "" and anyone_pregnant_notes <> "" Then err_msg = err_msg & "~!~" & "3 ^* Is anyone pregnant?##~##   - Since there are notes about pregnancy, indicate if anyone is pregnant."
    anyone_served_notes = trim(anyone_served_notes)
    If anyone_served_yn = "Yes" and anyone_served_notes = "" Then err_msg = err_msg & "~!~" & "3 ^* Has Anyone Served in the Military? - Notes##~##   - Since it is indicated that someone served in the military, add notes with additional details."
    If anyone_served_yn = "" and anyone_served_notes <> "" Then err_msg = err_msg & "~!~" & "3 ^* Has Anyone Served in the Military?##~##   - Since there are notes about military service, indicate if anyone has served in the military."

	If snap_status <> "INACTIVE" AND pwe_selection = "Select One..." Then err_msg = err_msg & "~!~" & "3 ^* Principal Wage Earner##~##   - Since we have SNAP to consider, you must indicate who the resident selects as PWE."
    If InStr(err_msg, "~!~3 ^*") Then page_3_errors = True

    err_selected_memb = ""
	For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_memb) = False Then
            pers_err = ""
            HH_MEMB_ARRAY(imig_status, the_memb) = trim(HH_MEMB_ARRAY(imig_status, the_memb))
            If HH_MEMB_ARRAY(requires_update, the_memb) Then pers_err = pers_err & "~!~" & "3 ^* Information Needed for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ":"
            If the_memb = 0 AND (HH_MEMB_ARRAY(id_verif, the_memb) = "" OR HH_MEMB_ARRAY(id_verif, the_memb) = "NO - No Ver Prvd") Then pers_err = pers_err & "~!~" & "3 ^* Identidty Verification##~##   - Identity is required for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ". Enter the ID information on file/received or indicate that it has been requested."

            If HH_MEMB_ARRAY(none_req_checkbox, the_memb) = unchecked Then
                If trim(HH_MEMB_ARRAY(ssn, the_memb)) = "" Then
                    If HH_MEMB_ARRAY(ssn_verif, the_memb) <> "A - SSN Applied For" and HH_MEMB_ARRAY(ssn_verif, the_memb) <> "N - Member Does Not Have SSN" Then
                        pers_err = pers_err & "~!~" & "3 ^* SSN##~##   - SSN is blank and should be requested now."
                    End If
                End If
                If HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - SSN Not Provided" Then
                    pers_err = pers_err & "~!~" & "3 ^* SSN##~##   - SSN Verification indicates not provided and should be requested now."
                End If
            End If

            If HH_MEMB_ARRAY(citizen, the_memb) = "No" and (HH_MEMB_ARRAY(none_req_checkbox, the_memb) = unchecked or the_memb = 0) Then
                If HH_MEMB_ARRAY(imig_status, the_memb) = "" Then pers_err = pers_err & "~!~" & "3 ^* IMIG Status##~##   - " & HH_MEMB_ARRAY(full_name_const, the_memb) & " is a non-citizen, discuss and record immigration status details."
                If HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "" or HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "?" Then pers_err = pers_err & "~!~" & "3 ^* Sponsor?##~##   - " & HH_MEMB_ARRAY(full_name_const, the_memb) & " is a non-citizen, you need to ask and record if this resident has a sponsor."
            End If
            prog_selected = False
            If HH_MEMB_ARRAY(snap_req_checkbox, the_memb) = checked Then prog_selected = True
            If HH_MEMB_ARRAY(cash_req_checkbox, the_memb) = checked Then prog_selected = True
            If HH_MEMB_ARRAY(emer_req_checkbox, the_memb) = checked Then prog_selected = True
            If prog_selected and HH_MEMB_ARRAY(none_req_checkbox, the_memb) = checked Then pers_err = pers_err & "~!~" & "3 ^* Program Request##~##   - Conflicting program request selections for " & HH_MEMB_ARRAY(full_name_const, the_memb) & ". If any program is checked, then 'NONE' cannot be checked."

            If pers_err <> "" Then
                If err_selected_memb = "" Then err_selected_memb = the_memb
                err_msg = err_msg & pers_err & "##~##"
            End If
    		' If HH_MEMB_ARRAY(intend_to_reside_in_mn, the_memb) = "" Then err_msg = err_msg & "~!~" & "3 ^* Intends to Reside in MN##~##   - Indicate if this resident (" & HH_MEMB_ARRAY(full_name_const, the_memb) & ") intends to reside in MN."
        End If
	Next
    If page_3_errors = True Then err_selected_memb = ""

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		list_err = False
		If FORM_QUESTION_ARRAY(quest).mandated = True Then
			list_err = True
			If FORM_QUESTION_ARRAY(quest).error_info = "school" AND (school_age_children_in_hh = False OR FORM_QUESTION_ARRAY(quest).interview_notes <> "") Then list_err = False
		End If

		If list_err = True Then err_msg = err_msg & "~!~" & FORM_QUESTION_ARRAY(quest).dialog_page_numb & " ^* " & FORM_QUESTION_ARRAY(quest).error_verbiage
	Next

	qual_memb_one = trim(qual_memb_one)
	qual_memb_two = trim(qual_memb_two)
	qual_memb_three = trim(qual_memb_three)
	qual_memb_four = trim(qual_memb_four)
	qual_memb_five = trim(qual_memb_five)
	If qual_question_one = "?" OR (qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type")) Then
		err_msg = err_msg & "~!~" & qual_numb & "^* Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
		If qual_question_one = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_one = "Yes" AND (qual_memb_one = "" OR qual_memb_one = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_two = "?" OR (qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type")) Then
		err_msg = err_msg & "~!~" & qual_numb & "^* Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
		If qual_question_two = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_two = "Yes" AND (qual_memb_two = "" OR qual_memb_two = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_three = "?" OR (qual_question_three = "Yes" AND (qual_memb_three = "" OR qual_memb_three = "Select or Type")) Then
		err_msg = err_msg & "~!~" & qual_numb & "^* Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
		If qual_question_three = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_three = "Yes" AND (qual_memb_three = "" OR qual_memb_three = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_four = "?" OR (qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type")) Then
		err_msg = err_msg & "~!~" & qual_numb & "^* Has anyone in your household been convicted of a drug felony in the past 10 years?"
		If qual_question_four = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_four = "Yes" AND (qual_memb_four = "" OR qual_memb_four = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If
	If qual_question_five = "?" OR (qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type")) Then
		err_msg = err_msg & "~!~" & qual_numb & "^* Is anyone in your household currently violating a condition of parole, probation or supervised release?"
		If qual_question_five = "?" Then err_msg = err_msg & "##~##   - Select 'Yes' or 'No' based on what the resident has entered on the CAF. If this is blank, ask the resident now."
		If qual_question_five = "Yes" AND (qual_memb_five = "" OR qual_memb_five = "Select or Type") Then err_msg = err_msg & "##~##   - Since this was answered 'Yes' you must indicate the person(s) who this 'Yes' applies to."
	End If

	If expedited_determination_needed = True Then
		If run_by_interview_team = False AND expedited_determination_completed = False Then err_msg = err_msg & "~!~" & exp_num & "^* Expedited##~##   - You must complete the process for the Expedited Determination. Press the 'EXPEDITED' button on the right and complete all steps."
		If run_by_interview_team = True Then
			If trim(exp_det_income) = "" Then err_msg = err_msg & "~!~" & exp_num & "^* Enter the amount of income in the application month for the Expedited Determination. If there is no income expected or received, enter a '0'"
			If trim(exp_det_assets) = "" Then err_msg = err_msg & "~!~" & exp_num & "^* Enter the amount of assets in the application month for the Expedited Determination. If there are no assets, enter a '0'"
			If trim(exp_det_housing) = "" Then err_msg = err_msg & "~!~" & exp_num & "^* Enter the expense for housing in the application month for the Expedited Determination. If there is no housing expense, enter a '0'"
			If trim(exp_det_income) <> "" and IsNumeric(exp_det_income)	= False Then err_msg = err_msg & "~!~" & exp_num & "^* The amount of income in the application month for the Expedited Determination must be entered as a number. This does not appear to be a number value, please review and reenter."
			If trim(exp_det_assets) <> "" and IsNumeric(exp_det_assets)	= False Then err_msg = err_msg & "~!~" & exp_num & "^* The amount of assets in the application month for the Expedited Determination must be entered as a number. This does not appear to be a number value, please review and reenter."
			If trim(exp_det_housing) <> "" and IsNumeric(exp_det_housing)	= False Then err_msg = err_msg & "~!~" & exp_num & "^* The expense for housing in the application month for the Expedited Determination must be entered as a number. This does not appear to be a number value, please review and reenter."
			utility_checked = false
			If heat_exp_checkbox = checked then utility_checked = true
			If ac_exp_checkbox = checked then utility_checked = true
			If electric_exp_checkbox = checked then utility_checked = true
			If phone_exp_checkbox = checked then utility_checked = true
			If utility_checked = true and none_exp_checkbox = checked Then err_msg = err_msg & "~!~" & exp_num & "^* A utility expense has been checked and also, NONE has been checked, review utilities for Expedited Determination."
			If utility_checked = false and none_exp_checkbox = unchecked Then err_msg = err_msg & "~!~" & exp_num & "^* No utility information was indicated for the Expedited Determination. If there are no utility expenses, check NONE."
		End If
	End If

	If err_msg = "" Then interview_questions_clear = TRUE

	If interview_questions_clear = TRUE Then
		' If current_listing = "11" Then tagline = ": CAF Last Page"
		'Both signatures - cannot be select or type or blank
		signature_detail = trim(signature_detail)
		second_signature_detail = trim(second_signature_detail)
		signature_person = trim(signature_person)
		second_signature_person = trim(second_signature_person)
		If signature_detail = "Select or Type" OR signature_detail = "" Then err_msg = err_msg & "~!~" & last_num & "^* Signature of Primary Adult##~##   - Indicate how the signature information has been received (or not received)."
		If second_signature_detail = "Select or Type" OR second_signature_detail = "" Then err_msg = err_msg & "~!~" & last_num & "^* Signature of Other Adult##~##   - Indicate how the second signature information has been received (or not received). If no second adult is on the case or the signature of the second adult is not required, select 'Not Required'."
		'If signatires are signed or verbal - then person and date must be completed
		If signature_detail = "Signature Completed" OR signature_detail  = "Accepted Verbally" Then
			If signature_person = "" AND signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & last_num & "^* Signature of Primary Adult - person##~##   - Since the signature was completed, indicate whose sigature it is."
		End If
		If second_signature_detail = "Signature Completed" OR second_signature_detail  = "Accepted Verbally" Then
			If second_signature_person = "" AND second_signature_person = "Select or Type" Then err_msg = err_msg & "~!~" & last_num & "^* Signature of Other Adult - person##~##   - Since the secondary adult signature was completed, indicate whose sigature it is."
		End If
		'Interview date must be a date and not in the future
		' If  Then err_msg = err_msg & "~!~" & "11^* FIELD##~##   - "
		If IsDate(interview_date) = False Then
			err_msg = err_msg & "~!~" & last_num & "^* Interview Date##~##   - Enter the date of the interview as a valid date."
		Else
			If DateDiff("d", date, interview_date) > 0 Then err_msg = err_msg & "~!~" & last_num & "^* Interview Date##~##   - The date of the interview cannot be in the future."
		End If

		If snap_status = "INACTIVE" AND case_is_expedited = True Then
			If pend_snap_on_case = "?" Then err_msg = err_msg & "~!~" & last_num & "^* SHOULD SNAP BE PENDED ##~##   - Since SNAP is not active on this case, review for possible program eligibility."
		End If
		IF family_cash_case = True OR adult_cash_case = True OR unknown_cash_pending = True Then
			If family_cash_case_yn = "?" Then
				err_msg = err_msg & "~!~" & last_num & "^* IS THIS A FAMILY CASH CASE ##~##   - Since this case has cash active or pending, indicate if this cash is MFIP/DWP."
			ElseIf family_cash_case_yn = "Yes" Then
				If absent_parent_yn = "?" Then err_msg = err_msg & "~!~" & last_num & "^* IS THERE AN ABPS ON THIS CASE ##~##   - Since this is a family cash case, indicate if there is an absent parent for any child on the case."
				If relative_caregiver_yn = "?" Then err_msg = err_msg & "~!~" & last_num & "^* IS THIS A RELATIVE CAREGIVER CASE ##~##   - Since this is a family cash case, indicate if this is a relative caregiver case."
			End If
 		End If


		If disc_no_phone_number = "EXISTS" Then err_msg = err_msg & "~!~" & discrep_num & "^* PHONE CONTACT Clarification ##~##   - Since no phone numbers were listed - confirm with the resident about phone contact and clarify."
		If disc_homeless_no_mail_addr = "EXISTS" Then err_msg = err_msg & "~!~" & discrep_num & "^* HOMELESS MAILING Clarification ##~##   - Since this case is listed as Homeless - confirm you have discussed mailing and responses."
		If disc_out_of_county = "EXISTS" Then err_msg = err_msg & "~!~" & discrep_num & "^* OUT OF COUNTY Clarification ##~##   - Since this case is indicated as being out of county - confirm you have explained case transfers."
		If disc_rent_amounts = "EXISTS" Then err_msg = err_msg & "~!~" & discrep_num & "^* HOUSING EXPENSE Clarification ##~##   - Since the amounts reported on the CAF for Housing Expense appear to have a discrepancy - clarify which is accurate."
		If disc_utility_amounts = "EXISTS" Then err_msg = err_msg & "~!~" & discrep_num & "^* UTILITY EXPENSE Clarification ##~##   - Since the amounts reported on the CAF for Utility Expense appear to have a discrepancy - clarify which is accurate."
	End If

	If EMER_on_CAF_checkbox = checked Then
		If trim(resident_emergency_yn) = "" or trim(emergency_type) = "" or emergency_type = "Select or Type" or trim(emergency_discussion) = "" or trim(emergency_amount) = "" or trim(emergency_deadline) = "" Then
			err_msg = err_msg & "~!~" & emer_numb & "^* The resident indicated they are applying for EMER. The EMER Q fields must all be filled out."
		End If
	End If

	If resident_emergency_yn = "Yes" Then
		If trim(emergency_type) = "" or emergency_type = "Select or Type" or trim(emergency_discussion) = "" or trim(emergency_amount) = "" or trim(emergency_deadline) = "" Then
			err_msg = err_msg & "~!~" & emer_numb & "^* You indicated the resident is experiencing an emergency. You must fill out all of the fields describing the emergency."
		End If
	End If
end function

function define_main_dialog()

	BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions   ---   Questions from " & CAF_form

        ' If CAF_form = "MNbenefits" Then Text 485, 5, 10, 10, "COVER LETTER"
        If CAF_form <> "MNbenefits" Then Text 485, 5, 75, 10, "---   DIALOGS   ---"
		Text 485, 22, 10, 10, "1"
		Text 485, 37, 10, 10, "2"
		Text 485, 52, 10, 10, "3"
		Text 485, 67, 10, 10, "4"
		num_pos = 82
		If last_page_of_questions => 5 Then
			Text 485, num_pos, 10, 10, "5"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 6 Then
			Text 485, num_pos, 10, 10, "6"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 7 Then
			Text 485, num_pos, 10, 10, "7"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 8 Then
			Text 485, num_pos, 10, 10, "8"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 9 Then
			Text 485, num_pos, 10, 10, "9"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 10 Then
			Text 485, num_pos, 10, 10, "10"
			num_pos = num_pos + 15
		End If
		If last_page_of_questions => 11 Then
			Text 485, num_pos, 10, 10, "11"
			num_pos = num_pos + 15
		End If

		qual_numb = show_qual
		show_qual = show_qual * 1
		Text 485, num_pos, 10, 10, qual_numb
		qual_pos = num_pos
		num_pos = num_pos + 15

		emer_numb = emergency_questions
		emergency_questions = emergency_questions * 1
		Text 485, num_pos, 10, 10, emer_numb
		emer_pos = num_pos
		num_pos = num_pos + 15

		running_num = last_page_of_questions + 3
		If discrepancies_exist = True Then
			discrep_num = running_num
			Text 485, num_pos, 10, 10, running_num
			discrep_pos = num_pos
			running_num = running_num + 1
			num_pos = num_pos + 15
		End If

		If expedited_determination_needed = True Then
			exp_num = running_num
			Text 485, num_pos, 10, 10, running_num
			exp_pos = num_pos
			running_num = running_num + 1
			num_pos = num_pos + 15
		End If

		last_num = running_num
		Text 485, num_pos, 10, 10, running_num
		last_pos = num_pos
		num_pos = num_pos + 15
	    ButtonGroup ButtonPressed

        If page_display = show_cover_letter and CAF_form = "MNbenefits" Then
            Text 490, 7, 60, 10, "COVER LETTER"
			Text 15, 15, 300, 10, "Before starting the interview questions, record the details from the MNBeneftis Cover Letter."
            y_pos = 30


            'IF EMER IS CHECKED
            If EMER_on_CAF_checkbox = checked or emer_verbal_request = "Yes" Then
                GroupBox 5, y_pos, 475, 50, "Since EMER is requested, the cover letter may have EMER information"
                y_pos = y_pos + 10
                Text 15, y_pos+5, 75, 10, "Emergency Type:"
                ComboBox 90, y_pos, 210, 25, "Select or Type"+chr(9)+"Eviction"+chr(9)+"Forced Move"+chr(9)+"Foreclosure"+chr(9)+"Utility Disconnect"+chr(9)+"Home Repairs"+chr(9)+"Property Taxes"+chr(9)+"Bus Ticket"+chr(9)+emergency_type, emergency_type
                Text 15, y_pos+25, 60, 10, "Comments:"
                EditBox 75, y_pos+20, 390, 15, emergency_discussion
                y_pos = y_pos + 50
            End If

            GB_y_pos = y_pos
            y_pos = y_pos + 15
            Text 15, y_pos, 130, 10, "Additional Application Comments:"
            EditBox 15, y_pos+10, 455, 15, additional_application_comments
            y_pos = y_pos + 35

            Text 5, y_pos, 475, 10, "----- Jobs and Self Employment Listed ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            y_pos = y_pos + 15
            mn_ben_job_quest = 8
            mn_ben_busi_quest = 9
            mn_ben_unea_quest = 13
            call FORM_QUESTION_ARRAY(mn_ben_job_quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, mn_ben_job_quest), TEMP_INFO_ARRAY(form_write_in_const, mn_ben_job_quest), TEMP_INFO_ARRAY(intv_notes_const, mn_ben_job_quest), TEMP_INFO_ARRAY(form_second_yn_const, mn_ben_job_quest), TEMP_INFO_ARRAY(form_second_ans_const, mn_ben_job_quest), "", False)
            call FORM_QUESTION_ARRAY(mn_ben_busi_quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, mn_ben_busi_quest), TEMP_INFO_ARRAY(form_write_in_const, mn_ben_busi_quest), TEMP_INFO_ARRAY(intv_notes_const, mn_ben_busi_quest), TEMP_INFO_ARRAY(form_second_yn_const, mn_ben_busi_quest), TEMP_INFO_ARRAY(form_second_ans_const, mn_ben_busi_quest), "", False)

            Text 5, y_pos-10, 475, 10, "----- Unearned Income Listed -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            y_pos = y_pos + 5
            call FORM_QUESTION_ARRAY(mn_ben_unea_quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, mn_ben_unea_quest), TEMP_INFO_ARRAY(form_write_in_const, mn_ben_unea_quest), TEMP_INFO_ARRAY(intv_notes_const, mn_ben_unea_quest), TEMP_INFO_ARRAY(form_second_yn_const, mn_ben_unea_quest), TEMP_INFO_ARRAY(form_second_ans_const, mn_ben_unea_quest), "", False)

            Text 5, y_pos-10, 475, 10, "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            y_pos = y_pos + 5

            Text 15, y_pos, 130, 10, "Additional Income Comments:"
            EditBox 15, y_pos+10, 455, 15, additional_income_comments
            y_pos = y_pos + 35
            Text 15, y_pos, 130, 10, "Interview Notes on Cover Letter Details:"
            EditBox 15, y_pos+10, 455, 15, cover_letter_interview_notes
            y_pos = y_pos + 35
            GroupBox 5, GB_y_pos, 475, y_pos - GB_y_pos, "APPLICATION COMMENTS AND INFORMATION"

        End If
        If page_display = show_pg_one_memb01_and_exp Then
			Text 497, 22, 60, 10, "INTVW / CAF 1"

			ComboBox 120, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_interview_with, who_are_we_completing_the_interview_with
			ComboBox 120, 30, 75, 45, "Select or Type"+chr(9)+"Phone"+chr(9)+"In Office"+chr(9)+how_are_we_completing_the_interview, how_are_we_completing_the_interview
			EditBox 120, 50, 50, 15, interview_date
			ComboBox 120, 70, 340, 45, "No Interpreter Used"+chr(9)+"Language Line Interpreter Used"+chr(9)+"Interpreter through Henn Co. OMS (Office of Multi-Cultural Services)"+chr(9)+"Interviewer speaks Resident Language"+chr(9)+interpreter_information, interpreter_information
			ComboBox 120, 90, 205, 45, "English"+chr(9)+"Somali"+chr(9)+"Spanish"+chr(9)+"Hmong"+chr(9)+"Russian"+chr(9)+"Oromo"+chr(9)+"Vietnamese"+chr(9)+interpreter_language, interpreter_language
            PushButton 330, 90, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
            EditBox 120, 110, 340, 15, arep_interview_id_information
			EditBox 10, 160, 450, 15, non_applicant_interview_info

			If expedited_screening_on_form = True Then
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
			End If

		    Text 10, 15, 110, 10, "Who are you interviewing with?"
			Text 65, 35, 55, 10, "Interview via"
			Text 65, 55, 55, 10, "Interview date"
			Text 30, 75, 85, 10, "Was an Interpreter Used?"
			Text 75, 95, 35, 10, "Language"
			Text 10, 115, 110, 10, "Detail AREP Identity Document"
			Text 120, 125, 300, 10, "- AREP ID is required if AREP applies on behalf of the resident."
			Text 120, 135, 300, 10, "- If no ID is required, this can be detailed here."
			Text 10, 150, 300, 10, "If interview is NOT with a Household Adult, explain relationship and add any details:"

			If expedited_screening_on_form = True Then
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
			End If

		ElseIf page_display = show_pg_one_address Then
			Text 504, 37, 60, 10, "CAF ADDR"
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
				DropListBox 175, 130, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", licensed_facility
				DropListBox 350, 130, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", meal_provided
  				Text 150, 155, 210, 10, residence_name_phone
				Text 70, 205, 305, 15, mail_addr_street_full
				Text 70, 225, 105, 15, mail_addr_city
				Text 205, 225, 110, 45, mail_addr_state
				Text 340, 225, 35, 15, mail_addr_zip
				Text 20, 280, 90, 15, phone_one_number
				Text 125, 280, 65, 45, phone_one_type
				Text 20, 300, 90, 15, phone_two_number
				Text 125, 300, 65, 45, phone_two_type
				Text 20, 320, 90, 15, phone_three_number
				Text 125, 320, 65, 45, phone_three_type
				Text 325, 255, 50, 15, address_change_date
				Text 255, 285, 120, 45, resi_addr_county
				PushButton 290, 340, 95, 15, "Update Information", update_information_btn
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
				DropListBox 175, 130, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", licensed_facility
				DropListBox 350, 130, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", meal_provided
				EditBox 150, 150, 230, 15, residence_name_phone
				EditBox 70, 200, 305, 15, mail_addr_street_full
				EditBox 70, 220, 105, 15, mail_addr_city
				DropListBox 205, 220, 110, 45, ""+chr(9)+state_list, mail_addr_state
				EditBox 340, 220, 35, 15, mail_addr_zip
				EditBox 20, 280, 90, 15, phone_one_number
				DropListBox 125, 280, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_one_type
				EditBox 20, 300, 90, 15, phone_two_number
				DropListBox 125, 300, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_two_type
				EditBox 20, 320, 90, 15, phone_three_number
				DropListBox 125, 320, 65, 45, "Select One..."+chr(9)+"C - Cell"+chr(9)+"H - Home"+chr(9)+"W - Work"+chr(9)+"M - Message"+chr(9)+"T - TTY/TDD", phone_three_type
				EditBox 325, 250, 50, 15, address_change_date
				ComboBox 255, 285, 120, 45, county_list+chr(9)+resi_addr_county, resi_addr_county
				PushButton 290, 340, 95, 15, "Save Information", save_information_btn
			End If

			Text 255, 305, 125, 10, "Send Updates via TEXT MESSAGE:"
			DropListBox 380, 300, 50, 15, "?"+chr(9)+"No"+chr(9)+"Yes", send_text
			Text 290, 325, 90, 10, "Send Updates via EMAIL:"
			DropListBox 380, 320, 50, 15, "?"+chr(9)+"No"+chr(9)+"Yes", send_email

			PushButton 325, 185, 50, 10, "CLEAR", clear_mail_addr_btn
			PushButton 205, 280, 35, 10, "CLEAR", clear_phone_one_btn
			PushButton 205, 300, 35, 10, "CLEAR", clear_phone_two_btn
			PushButton 205, 320, 35, 10, "CLEAR", clear_phone_three_btn

			Text 10, 10, 450, 10, "Review the Address informaiton known with the resident. If it needs updating, press the 'Update Information' button to make changes:"
			GroupBox 10, 35, 375, 135, "Residence Address"
			Text 20, 55, 45, 10, "House/Street"
			Text 45, 75, 20, 10, "City"
			Text 185, 75, 20, 10, "State"
			Text 325, 75, 15, 10, "Zip"
			Text 20, 95, 100, 10, "Do you live on a Reservation?"
			Text 180, 95, 60, 10, "If yes, which one?"
			Text 20, 115, 100, 10, "Resident Indicates Homeless:"
			Text 185, 115, 60, 10, "Living Situation?"
			Text 20, 135, 155, 10, "Are you currently residing in a licensed facility?"
			Text 210, 135, 140, 10, "Does the place you reside provide meals?"
			Text 20, 155, 130, 10, "Name and phone number of residence: "
			GroupBox 10, 175, 375, 70, "Mailing Address"
			Text 20, 205, 45, 10, "House/Street"
			Text 45, 225, 20, 10, "City"
			Text 185, 225, 20, 10, "State"
			Text 325, 225, 15, 10, "Zip"
			GroupBox 10, 250, 235, 90, "Phone Number"
			Text 20, 265, 50, 10, "Number"
			Text 125, 265, 25, 10, "Type"
			Text 255, 255, 60, 10, "Date of Change:"
			Text 255, 275, 75, 10, "County of Residence:"

		ElseIf page_display = show_pg_memb_list Then
			Text 504, 52, 60, 10, "CAF MEMBs"
            allow_expand = False
            If UBound(HH_MEMB_ARRAY, 2) < 10 Then allow_expand = True

            If update_pers = FALSE Then
                EditBox 800, 500, 50, 15, dummy_editbox_to_capture_focus
                y_pos = 20
                If HH_arrived_date <> "" Then
                    Text 20, y_pos, 200, 10, "*** Household Arrived in Minnesota on " & HH_arrived_date & " from " & HH_arrived_place
                    y_pos = y_pos + 15
                End If
                For the_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
                    progs = ""
                    If HH_MEMB_ARRAY(snap_req_checkbox, the_membs) = checked Then progs = progs & "SNAP "
                    If HH_MEMB_ARRAY(cash_req_checkbox, the_membs) = checked Then progs = progs & "CASH "
                    If HH_MEMB_ARRAY(emer_req_checkbox, the_membs) = checked Then progs = progs & "EMER "
                    progs = replace(trim(progs), " ", " - ")
                    If progs = "" Then progs = " (none)"

                    ' MEMB_requires_update = False
                    member_info_string = ""
                    If HH_MEMB_ARRAY(rel_to_applcnt, the_membs) = "01 Self" and (HH_MEMB_ARRAY(id_verif, the_membs) = "__" or HH_MEMB_ARRAY(id_verif, the_membs) = "NO - No Ver Prvd") Then
                        ' MEMB_requires_update = True
                        member_info_string = member_info_string & "ID Missing; "
                    End If
                    If HH_MEMB_ARRAY(spouse_ref, the_membs) <> "" Then member_info_string = member_info_string & "Spouse: M " & HH_MEMB_ARRAY(spouse_ref, the_membs) & "; "
                    If trim(HH_MEMB_ARRAY(ssn, the_memb)) = "" Then
                        If HH_MEMB_ARRAY(ssn_verif, the_memb) <> "A - SSN Applied For" and HH_MEMB_ARRAY(ssn_verif, the_memb) <> "N - Member Does Not Have SSN" Then
                            member_info_string = member_info_string & "SSN Missing; "
                        End If
                    End If
                    If HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - SSN Not Provided" Then
                        member_info_string = member_info_string & "SSN Not Provided; "
                    End If
                    If HH_MEMB_ARRAY(citizen, the_membs) = "No" and trim(progs) <> "(none)" Then
                        member_info_string = member_info_string & "*** Non-Citizen ***; "
                        If trim(HH_MEMB_ARRAY(imig_status, the_membs)) = "" Then
                            ' MEMB_requires_update = True
                            If HH_MEMB_ARRAY(clt_has_sponsor, the_membs) = "?" or HH_MEMB_ARRAY(clt_has_sponsor, the_membs) = "" Then
                                member_info_string = member_info_string & "Imig Status and Sponsor Info Missing; "
                            Else
                                member_info_string = member_info_string & "Imig Status Missing; "
                            End If
                        Else
                            member_info_string = member_info_string & "Imig Status: " & HH_MEMB_ARRAY(imig_status, the_membs) & "; "
                            If HH_MEMB_ARRAY(clt_has_sponsor, the_membs) = "?" or HH_MEMB_ARRAY(clt_has_sponsor, the_membs) = "" Then
                                ' MEMB_requires_update = True
                                member_info_string = member_info_string & "Sponsor Info Missing; "
                            End If
                        End If
                    End If
                    If HH_MEMB_ARRAY(in_mn_12_mo, the_membs) = "No" and HH_arrived_place = "" and HH_arrived_date = "" Then
                        If HH_MEMB_ARRAY(former_state, the_membs) = "NB" Then member_info_string = member_info_string & "Born on " & HH_MEMB_ARRAY(mn_entry_date, the_membs) & "; "
                        If HH_MEMB_ARRAY(former_state, the_membs) <> "NB" Then member_info_string = member_info_string & "MN Entry: " & HH_MEMB_ARRAY(mn_entry_date, the_membs) & " from " & HH_MEMB_ARRAY(former_state, the_membs) & "; "
                    End If
                    If HH_MEMB_ARRAY(interpreter, the_membs) = "Yes" and (HH_MEMB_ARRAY(age, the_membs) > 17 OR the_membs = 0) Then member_info_string = member_info_string & "Interpreter Needed; "

                    If trim(progs) <> "(none)" and (HH_MEMB_ARRAY(age, the_membs) > 17 OR the_membs = 0) Then
                        If left(HH_MEMB_ARRAY(spoken_lang, the_membs), 2) <> "99" and len(HH_MEMB_ARRAY(spoken_lang, the_membs)) > 5 Then member_info_string = member_info_string & "Language: " & right(HH_MEMB_ARRAY(spoken_lang, the_membs), len(HH_MEMB_ARRAY(spoken_lang, the_membs))-5) & "; "
                        If len(HH_MEMB_ARRAY(spoken_lang, the_membs)) < 6 Then member_info_string = member_info_string & "Spoken Lang Unknown; "
                    End If
                    If trim(progs) <> "(none)" Then
                        If HH_MEMB_ARRAY(race, the_membs) = "Unable To Determine" Then member_info_string = member_info_string & "Race Undetermined; "
                    End If
                    If HH_MEMB_ARRAY(alias_yn, the_membs) = "Yes" Then member_info_string = member_info_string & "Alias Exists; "

                    ' If HH_MEMB_ARRAY(none_req_checkbox, the_membs) = checked Then MEMB_requires_update = False
                    If HH_MEMB_ARRAY(requires_update, the_membs) Then member_info_string = "*** UPDATE! - " & member_info_string

                    member_info_string = trim(member_info_string)
                    If right(member_info_string, 1) = ";" Then member_info_string = left(member_info_string, len(member_info_string)-1)

                    If len(HH_MEMB_ARRAY(full_name_const, the_membs)) > 25 Then
                        If HH_MEMB_ARRAY(pers_in_maxis, the_membs) Then     Text 15,  y_pos, 225, 10, "M " & HH_MEMB_ARRAY(ref_number, the_membs) & "   -   " & HH_MEMB_ARRAY(first_name_const, the_membs)
                        If NOT HH_MEMB_ARRAY(pers_in_maxis, the_membs) Then Text 15,  y_pos, 225, 10, HH_MEMB_ARRAY(first_name_const, the_membs)
                        Text        45,  y_pos+10, 225, 10, HH_MEMB_ARRAY(last_name_const, the_membs)
                    Else
                        If HH_MEMB_ARRAY(pers_in_maxis, the_membs) Then     Text 15,  y_pos, 225, 10, "M " & HH_MEMB_ARRAY(ref_number, the_membs) & "   -   " & HH_MEMB_ARRAY(full_name_const, the_membs)
                        If NOT HH_MEMB_ARRAY(pers_in_maxis, the_membs) Then Text 15,  y_pos, 225, 10, HH_MEMB_ARRAY(full_name_const, the_membs)
                    End If
                    If len(HH_MEMB_ARRAY(full_name_const, the_membs)) < 26 Then Text 45,  y_pos+10, 50, 10, "Age: " & HH_MEMB_ARRAY(age, the_membs)
                    If len(HH_MEMB_ARRAY(full_name_const, the_membs)) > 25 Then Text 125,  y_pos, 50, 10, "Age: " & HH_MEMB_ARRAY(age, the_membs)
                    If the_membs <> 0 Then Text        185,  y_pos+10, 100, 10, "Rel to 01: " & right(HH_MEMB_ARRAY(rel_to_applcnt, the_membs), len(HH_MEMB_ARRAY(rel_to_applcnt, the_membs))-3)

                    Text        180, y_pos, 75,  10, progs

                    PushButton  415, y_pos, 60,  10, "UPDATE M " & HH_MEMB_ARRAY(ref_number, the_membs), HH_MEMB_ARRAY(button_one, the_membs)
                    If allow_expand and len(member_info_string) > 130 Then
                        Text        270, y_pos, 145, 35, member_info_string
                        y_pos = y_pos + 15
                    ElseIf allow_expand and len(member_info_string) > 80 Then
                        Text        270, y_pos, 145, 30, member_info_string
                        y_pos = y_pos + 10
                    Else
                        Text        270, y_pos, 145, 20, member_info_string
                    End If

                    y_pos = y_pos + 20
                Next
                GroupBox 5, 5, 470, y_pos, "All Household MEMBERs"
                Text 180, 5, 50, 10, "REQUESTING:"
                Text 270, 5, 105, 10, "INFORMATION NEEDED:"
                y_pos = y_pos + 10

                ' Text 10, y_pos+5, 155, 10, "Are all people in the Houshold Listed above?"
                ' DropListBox 160, y_pos, 35, 35, ""+chr(9)+"Yes"+chr(9)+"No", all_members_listed_yn
                Text 10, y_pos+5, 60, 10, "HH Comp Notes:"
                EditBox 70, y_pos, 350, 15, all_members_listed_notes
                PushButton 425, y_pos, 50, 15, "Add Person", add_person_btn
                ' PushButton 370, y_pos, 105, 15, "Enter Date Person Left", remove_person_btn
                y_pos = y_pos + 20

                Text 10, y_pos+5, 140, 10, "Do ALL Members intend to reside in MN?"
                DropListBox 150, y_pos, 35, 45, ""+chr(9)+"Yes"+chr(9)+"No", all_members_in_MN_yn
                Text 190, y_pos+5, 25, 10, "Notes:"
                EditBox 215, y_pos, 260, 15, all_members_in_MN_notes
                y_pos = y_pos + 25

                Text 10, y_pos+5, 70, 10, "Is Anyone Pregnant?"
                DropListBox 80, y_pos, 35, 45, ""+chr(9)+"Yes"+chr(9)+"No", anyone_pregnant_yn
                Text 120, y_pos+5, 25, 10, "Notes:"
                EditBox 145, y_pos, 330, 15, anyone_pregnant_notes
                y_pos = y_pos + 25

                Text 10, y_pos+5, 115, 10, "Has Anyone Served in the Military?"
                DropListBox 125, y_pos, 35, 45, ""+chr(9)+"Yes"+chr(9)+"No", anyone_served_yn
                Text 165, y_pos+5, 25, 10, "Notes:"
                EditBox 190, y_pos, 285, 15, anyone_served_notes
                y_pos = y_pos + 25

                Text 10, y_pos+5, 75, 10, "Principal Wage Earner:"
                DropListBox 85, y_pos, 130, 45, pick_a_client, pwe_selection
                y_pos = y_pos + 20

			End If
			If update_pers = TRUE Then
                EditBox 25, 30, 90, 15, HH_MEMB_ARRAY(last_name_const, selected_memb)
                EditBox 120, 30, 75, 15, HH_MEMB_ARRAY(first_name_const, selected_memb)
                EditBox 200, 30, 50, 15, HH_MEMB_ARRAY(mid_initial, selected_memb)
                EditBox 255, 30, 105, 15, HH_MEMB_ARRAY(other_names, selected_memb)
                EditBox 370, 30, 60, 15, HH_MEMB_ARRAY(date_of_birth, selected_memb)
                If HH_MEMB_ARRAY(ssn_verif, selected_memb) = "V - SSN Verified via Interface" Then Text 25, 65, 50, 10, HH_MEMB_ARRAY(ssn, selected_memb)
                If HH_MEMB_ARRAY(ssn_verif, selected_memb) <> "V - SSN Verified via Interface" Then EditBox 25, 60, 50, 15, HH_MEMB_ARRAY(ssn, selected_memb)
                DropListBox 75, 60, 110, 15, ssn_verif_list, HH_MEMB_ARRAY(ssn_verif, selected_memb)
                DropListBox 190, 60, 40, 45, "Male"+chr(9)+"Female", HH_MEMB_ARRAY(gender, selected_memb)
                DropListBox 235, 60, 75, 45, memb_panel_relationship_list, HH_MEMB_ARRAY(rel_to_applcnt, selected_memb)
                DropListBox 325, 60, 105, 45, marital_status_list, HH_MEMB_ARRAY(marital_status, selected_memb)
                EditBox 25, 90, 110, 15, HH_MEMB_ARRAY(last_grade_completed, selected_memb)
                EditBox 140, 90, 70, 15, HH_MEMB_ARRAY(mn_entry_date, selected_memb)
                EditBox 215, 90, 135, 15, HH_MEMB_ARRAY(former_state, selected_memb)
                DropListBox 355, 90, 75, 45, "Yes"+chr(9)+"No", HH_MEMB_ARRAY(citizen, selected_memb)
                DropListBox 25, 120, 60, 45, "No"+chr(9)+"Yes", HH_MEMB_ARRAY(interpreter, selected_memb)
                EditBox 95, 120, 120, 15, HH_MEMB_ARRAY(spoken_lang, selected_memb)
                EditBox 95, 150, 120, 15, HH_MEMB_ARRAY(written_lang, selected_memb)
                DropListBox 285, 130, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(ethnicity_yn, selected_memb)
                DropListBox 25, 170, 110, 45, ""+chr(9)+id_droplist_info, HH_MEMB_ARRAY(id_verif, selected_memb)
                PushButton 340, 210, 95, 15, "Save Information", save_information_btn
                CheckBox 285, 155, 30, 10, "Asian", HH_MEMB_ARRAY(race_a_checkbox, selected_memb)
                CheckBox 285, 165, 30, 10, "Black", HH_MEMB_ARRAY(race_b_checkbox, selected_memb)
                CheckBox 285, 175, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(race_n_checkbox, selected_memb)
                CheckBox 285, 185, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(race_p_checkbox, selected_memb)
                CheckBox 285, 195, 130, 10, "White", HH_MEMB_ARRAY(race_w_checkbox, selected_memb)
                CheckBox 25, 195, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(snap_req_checkbox, selected_memb)
                CheckBox 80, 195, 65, 10, "Cash programs", HH_MEMB_ARRAY(cash_req_checkbox, selected_memb)
                CheckBox 150, 195, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(emer_req_checkbox, selected_memb)
                CheckBox 235, 195, 30, 10, "NONE", HH_MEMB_ARRAY(none_req_checkbox, selected_memb)
                EditBox 110, 210, 95, 15, HH_MEMB_ARRAY(remo_info_const, known_membs)
                If selected_memb = 0 Then
                    DropListBox 25, 250, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb)
                Else
                    DropListBox 25, 250, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Not in HH", HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb)
                End If
                EditBox 110, 250, 205, 15, HH_MEMB_ARRAY(imig_status, selected_memb)
                DropListBox 320, 250, 55, 45, "?"+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(clt_has_sponsor, selected_memb)
                DropListBox 25, 280, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(client_verification, selected_memb)
                EditBox 110, 280, 320, 15, HH_MEMB_ARRAY(client_verification_details, selected_memb)
                EditBox 25, 310, 405, 15, HH_MEMB_ARRAY(client_notes, selected_memb)
                If HH_MEMB_ARRAY(ref_number, selected_memb) = "" Then
                    GroupBox 20, 10, 415, 200, "Person " & selected_memb+1 & " " & HH_MEMB_ARRAY(full_name_const, selected_memb)
                    GroupBox 20, 230, 415, 100, "Person " & selected_memb+1 & " " & HH_MEMB_ARRAY(full_name_const, selected_memb) & "  ---  Interview Questions"
                Else
                    GroupBox 20, 10, 415, 200, "MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb) & " - " & HH_MEMB_ARRAY(full_name_const, selected_memb)
                    GroupBox 20, 230, 415, 100, "MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb) & " - " & HH_MEMB_ARRAY(full_name_const, selected_memb) & "  ---  Interview Questions"

                End If

                If HH_MEMB_ARRAY(pers_in_maxis, selected_memb) = False Then PushButton 330, 15, 105, 15, "Remove Member from Script", HH_MEMB_ARRAY(button_two, selected_memb)
                Text 25, 20, 50, 10, "Last Name"
                Text 120, 20, 50, 10, "First Name"
                Text 200, 20, 50, 10, "Middle Name"
                Text 255, 20, 50, 10, "Other Names"
                Text 370, 20, 45, 10, "Date of Birth"
                Text 25, 50, 100, 10, "Social Security Number"
                ' Text 25, 50, 55, 10, "SSN"
                Text 190, 50, 45, 10, "Gender"
                Text 235, 50, 90, 10, "Relationship to MEMB 01"
                Text 325, 50, 50, 10, "Marital Status"
                Text 25, 80, 75, 10, "Last Grade Completed"
                Text 140, 80, 55, 10, "Moved to MN on"
                Text 215, 80, 65, 10, "Moved to MN from"
                Text 355, 80, 75, 10, "US Citizen or National"
                Text 25, 110, 40, 10, "Interpreter?"
                Text 95, 110, 95, 10, "Preferred Spoken Language"
                Text 95, 140, 95, 10, "Preferred Written Language"
                Text 25, 160, 65, 10, "Identity Verification"
                GroupBox 280, 110, 155, 100, "Demographics"
                Text 285, 120, 35, 10, "Hispanic?"
                Text 285, 145, 50, 10, "Race"
                Text 25, 185, 145, 10, "Which programs is this person requesting?"
                Text 25, 215, 80, 10, "Date Member Left HH"
                Text 25, 240, 80, 10, "Intends to reside in MN"
                Text 110, 240, 65, 10, "Immigration Status"
                Text 320, 240, 50, 10, "Sponsor?"
                Text 25, 270, 50, 10, "Verification"
                Text 110, 270, 65, 10, "Verification Details"
                Text 25, 300, 50, 10, "Notes:"
                If HH_MEMB_ARRAY(requires_update, selected_memb) Then
                    Text 350, 5, 100, 10, "*** UPDATE REQUIRED ***"
                    Text 20, 330, 100, 10, "*** UPDATE REQUIRED ***"
                End If
                y_pos = 340
                If selected_memb = 0 AND (HH_MEMB_ARRAY(id_verif, selected_memb) = "" OR HH_MEMB_ARRAY(id_verif, selected_memb) = "NO - No Ver Prvd") Then
                    Text 25, y_pos, 400, 10, " - NO ID Verification for MEMBER 01."
                    y_pos = y_pos + 10
                End If
                If (trim(HH_MEMB_ARRAY(ssn, the_memb)) = "" and HH_MEMB_ARRAY(ssn_verif, the_memb) <> "A - SSN Applied For" and HH_MEMB_ARRAY(ssn_verif, the_memb) <> "N - Member Does Not Have SSN") or HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - SSN Not Provided" Then
                    Text 25, y_pos, 400, 10, " - SSN Information Missing"
                    y_pos = y_pos + 10
                End If
                If HH_MEMB_ARRAY(citizen, selected_memb) = "No" Then
                    Text 25, y_pos, 400, 10, " - NON-CITIZEN: Immigrations Status and Sponsor Info Needed."
                    y_pos = y_pos + 10
                End If
                If y_pos <> 340 Then Text 20, 330, 100, 10, "MEMBER Notes:"
			End If

		ElseIf page_display = show_qual Then
			Text 500, qual_pos, 60, 10, "CAF QUAL Q"

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
		ElseIf page_display = discrepancy_questions Then
			Text 504, discrep_pos, 60, 10, "Clarifications"

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
				ComboBox 120, y_pos + 15, 335, 45, "Select or Type"+chr(9)+"Phone paid by Government Free Phone Program with no expense."+chr(9)+"Phone is paid by someone out of the home, billed directly to them."+chr(9)+"Phone is a community line available for messages only."+chr(9)+"Phone is a community line in the building/residence the resident stays at."+chr(9)+"Resident uses free phone via internet program and pays no phone or internet bill"+chr(9)+disc_yes_phone_no_expense_confirmation, disc_yes_phone_no_expense_confirmation
				y_pos = y_pos + 40
			End If
			If disc_no_phone_yes_expense = "EXISTS" OR disc_no_phone_yes_expense = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 35, "No Phone Number Listed, Phone Expense Indicated"
				Text 20, y_pos + 20, 165, 10, "Clarify a phone number or explain expense:"
				ComboBox 185, y_pos + 15, 270, 45, "Select or Type"+chr(9)+"Paying phone for someone outside the home."+chr(9)+"Lost phone, number is changing."+chr(9)+"Getting a new number."+chr(9)+disc_no_phone_yes_expense_confirmation, disc_no_phone_yes_expense_confirmation
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
				Text 20, y_pos + 30, 400, 10, "Question on Housing Expense: " & rent_summary

				Text 20, y_pos + 50, 110, 10, "Confirm Housing Expense Detail: "
				ComboBox 125, y_pos + 45, 330, 45, "Select or Type"+chr(9)+"Houshold DOES have Housing Expense"+chr(9)+"Household has NO Housing expense"+chr(9)+"Houshold has an ongoing Housing Expense but NONE in the Application month"+chr(9)+"Houshold has Housing Expense in the application months but NONE ongoing"+chr(9)+disc_rent_amounts_confirmation, disc_rent_amounts_confirmation
				y_pos = y_pos + 70
			End If
			If disc_utility_amounts = "EXISTS" OR disc_utility_amounts = "RESOLVED" Then
				GroupBox 10, y_pos, 455, 65, "CAF Answers for Utility Expense do not Match, Review and Clarify"
				Text 20, y_pos + 15, 400, 10, "CAF Page 1 Utility Expense: " & disc_utility_caf_1_summary
				Text 20, y_pos + 30, 400, 10, "Question on Utility Expense: " & utility_summary

				Text 20, y_pos + 50, 110, 10, "Confirm Utility Expense Detail: "
				ComboBox 125, y_pos + 45, 330, 45, "Select or Type"+chr(9)+"Household pays for Heat"+chr(9)+"Household pays for AC"+chr(9)+"Houshold pays Electricity which INCLUDES AC"+chr(9)+"Houshold pays Electricity which INCLUDES Heat"+chr(9)+"Houshold pays Electricity which INCLUDES AC and Heat"+chr(9)+"Houshold pays Electricity, but this does not include Heat or AC"+chr(9)+"Houshold pays Electricity and Phone"+chr(9)+"Houshold pays Phone Only"+chr(9)+"Houshold pays NO Utility Expenses"+chr(9)+disc_utility_amounts_confirmation, disc_utility_amounts_confirmation
				y_pos = y_pos + 70
			End If

		ElseIf page_display = expedited_determination Then
			If expedited_viewed = False Then Call set_initial_exp_simplified
			expedited_viewed = True
			expedited_determination_completed = True

			Text 505, exp_pos, 60, 10, "EXPEDITED"

			GroupBox 5, 10, 475, 325, "Expedited Detail"
			Text 15, 25, 450, 20, "SNAP benefits can be issued quickly in certain circumstances. Discuss income, assets, and expenses with resident to document the correct amounts for the month of application. These do not have to be exact but should be the most reasonable estimate."

			Text 15, 50, 290, 10, "How much income was received (or will be received) in the application month (MM/YY)?"
			EditBox 310, 45, 50, 15, exp_det_income
			y_pos = 60

			For quest = 0 to UBound(FORM_QUESTION_ARRAY)
				If InStr(FORM_QUESTION_ARRAY(quest).dialog_phrasing, "self-employed") <> 0 AND FORM_QUESTION_ARRAY(quest).info_type = "single-detail" Then
					Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
					Text 30, y_pos+10, 450, 10, "CAF Answer: " & FORM_QUESTION_ARRAY(quest).caf_answer
					Text 30, y_pos+20, 450, 10, FORM_QUESTION_ARRAY(quest).sub_phrase & ": " & FORM_QUESTION_ARRAY(quest).sub_answer
					y_pos = y_pos + 30
					If FORM_QUESTION_ARRAY(quest).write_in_info <> "" Then
						Text 30, y_pos, 450, 10, "Write-In: " & FORM_QUESTION_ARRAY(quest).write_in_info
						y_pos = y_pos + 10
					End If
					If FORM_QUESTION_ARRAY(quest).interview_notes <> "" Then
						Text 30, y_pos, 450, 10, "Interview Notes: " & FORM_QUESTION_ARRAY(quest).interview_notes
						y_pos = y_pos + 10
					End If
				End If
				If FORM_QUESTION_ARRAY(quest).detail_array_exists = true Then
					first_array = True
					for each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
						If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then
							If FORM_QUESTION_ARRAY(quest).detail_business(each_item) <> "" or FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then
								If first_array = True Then
									Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
									y_pos = y_pos + 10
									first_array = False
								End If
								Text 25, y_pos, 450, 10, "Employer: " & FORM_QUESTION_ARRAY(quest).detail_business(each_item) & "  - Employee: " & FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & "   - Gross Monthly Earnings: $ " & FORM_QUESTION_ARRAY(quest).detail_monthly_amount(each_item)
								y_pos = y_pos + 10
								If trim(FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item)) <> "" OR trim(FORM_QUESTION_ARRAY(quest).detail_hours_per_week(each_item)) <> "" Then
									Text 30, y_pos, 450, 10, "Hourly Wage: " & FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item) & " - Hours per Week: " & FORM_QUESTION_ARRAY(quest).detail_hours_per_week(each_item)
									y_pos = y_pos + 10
								End If
							End If
						ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then
							If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" or FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then
								If first_array = True Then
									Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
									y_pos = y_pos + 10
									first_array = False
								End If
								Text 25, y_pos, 450, 10, "Name: " & FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & "  - Type: " & FORM_QUESTION_ARRAY(quest).detail_type(each_item) & "   - Start Date: $ " & FORM_QUESTION_ARRAY(quest).detail_date(each_item) & "   - Amount: $ " & FORM_QUESTION_ARRAY(quest).detail_amount(each_item) & "   - Freq.: " & FORM_QUESTION_ARRAY(quest).detail_frequency(each_item)
								y_pos = y_pos + 10
							End If
						End If
					next
				End If
				'Condense UNEA information so information does not extend past dialog
				x_pos = 30
				yes_answer_count = 0
				If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then
					Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
					y_pos = y_pos + 10
					If FORM_QUESTION_ARRAY(quest).answer_is_array = True Then
						For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
							If FORM_QUESTION_ARRAY(quest).item_ans_list(each_caf_unea) = "Yes" Then
								Text x_pos, y_pos, 140, 10, FORM_QUESTION_ARRAY(quest).item_info_list(each_caf_unea) & " - Yes" & " - " & FORM_QUESTION_ARRAY(quest).item_detail_list(each_caf_unea)
								yes_answer_count = yes_answer_count + 1
								x_pos = x_pos + 140
								If x_pos > 310 Then
									x_pos = 30
									y_pos = y_pos + 10
								End If
							End If
						Next
						If yes_answer_count = 1 Or yes_answer_count = 2 Or yes_answer_count = 4 OR yes_answer_count = 5 or yes_answer_count = 7 or  yes_answer_count = 8 then y_pos = y_pos + 10
					End If
					If FORM_QUESTION_ARRAY(quest).write_in_info <> "" Then
						Text 30, y_pos, 400, 10, "Write-In: " & FORM_QUESTION_ARRAY(quest).write_in_info
						y_pos = y_pos + 10
					End If
					If FORM_QUESTION_ARRAY(quest).interview_notes <> "" Then
						Text 30, y_pos, 400, 10, "Interview Notes: " & FORM_QUESTION_ARRAY(quest).interview_notes
						y_pos = y_pos + 10
					End If
				End If
			Next
			If y_pos = 60 Then y_pos = y_pos + 5
			y_pos = y_pos + 5

			Text 15, y_pos, 330, 10, "How much does the household have in assets (accounts and cash) in the application month (MM/YY)?"
			EditBox 350, y_pos-5, 50, 15, exp_det_assets
			y_pos = y_pos + 10
			orig_y_pos = y_pos

			For quest = 0 to UBound(FORM_QUESTION_ARRAY)
				If FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then
					first_array = True
					If FORM_QUESTION_ARRAY(quest).detail_array_exists = true Then
						for each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
							If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" or FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then
								If first_array = True Then
									Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
									y_pos = y_pos + 10
									first_array = False
								End If
								Text 25, y_pos, 450, 10, "Owner: " & FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & "  - Type: " & FORM_QUESTION_ARRAY(quest).detail_type(each_item) & "  - Value: $ " & FORM_QUESTION_ARRAY(quest).detail_value(each_item)
								y_pos = y_pos + 10
							End If
						next
					End If
				End If
				If FORM_QUESTION_ARRAY(quest).info_type = "asset" Then
					Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
					y_pos = y_pos + 10
					If FORM_QUESTION_ARRAY(quest).answer_is_array = True Then
						For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
							If FORM_QUESTION_ARRAY(quest).item_ans_list(each_caf_unea) = "Yes" Then
								Text 30, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).item_info_list(each_caf_unea) & " - Yes"
								y_pos = y_pos + 10
							End If
						Next
					End If
					If FORM_QUESTION_ARRAY(quest).write_in_info <> "" Then
						Text 30, y_pos, 450, 10, "Write-In: " & FORM_QUESTION_ARRAY(quest).write_in_info
						y_pos = y_pos + 10
					End If
					If FORM_QUESTION_ARRAY(quest).interview_notes <> "" Then
						Text 30, y_pos, 450, 10, "Interview Notes: " & FORM_QUESTION_ARRAY(quest).interview_notes
						y_pos = y_pos + 10
					End If
				End If
			Next
			If y_pos = orig_y_pos Then y_pos = y_pos + 5
			y_pos = y_pos + 5

			Text 15, y_pos, 305, 10, "How much does the household pay in housing expenses in the application month (MM/YY)?"
			EditBox 320, y_pos-5, 50, 15, exp_det_housing
			y_pos = y_pos + 15

			Text 15, y_pos, 315, 10, "Which type of utilities is the household responsible to pay in the application month (MM/YY)?"
			y_pos = y_pos + 10
			CheckBox 25, y_pos, 45, 10, "Heat", heat_exp_checkbox
			CheckBox 90, y_pos, 70, 10, "Air Conditioning", ac_exp_checkbox
			CheckBox 175, y_pos, 45, 10, "Electric", electric_exp_checkbox
			CheckBox 240, y_pos, 55, 10, "Telephone", phone_exp_checkbox
			CheckBox 300, y_pos, 55, 10, "NONE", none_exp_checkbox
			y_pos = y_pos + 15

			For quest = 0 to UBound(FORM_QUESTION_ARRAY)
				If FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then
					Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
					y_pos = y_pos + 10
					If FORM_QUESTION_ARRAY(quest).detail_array_exists = true Then
						for each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
							If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then
								Text 25, y_pos, 395, 10, "Type: " & FORM_QUESTION_ARRAY(quest).detail_type(each_item) & "  - Amount: $ " & FORM_QUESTION_ARRAY(quest).detail_amount(each_item) & "  - Frequency: " & FORM_QUESTION_ARRAY(quest).detail_frequency(each_item)
								y_pos = y_pos + 10
							End If
						next
					End If
				End If
				If FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then
					housing_info_txt = "Housing Payment: $ " & FORM_QUESTION_ARRAY(quest).housing_payment
					If FORM_QUESTION_ARRAY(quest).heat_air_checkbox = unchecked Then housing_info_txt = housing_info_txt & "  -   [ ] Heat/AC "
					If FORM_QUESTION_ARRAY(quest).heat_air_checkbox = checked Then housing_info_txt = housing_info_txt & "  -   [X] Heat/AC "
					If FORM_QUESTION_ARRAY(quest).electric_checkbox = unchecked Then housing_info_txt = housing_info_txt & "   [ ] Electric "
					If FORM_QUESTION_ARRAY(quest).electric_checkbox = checked Then housing_info_txt = housing_info_txt & "   [X] Electric "
					If FORM_QUESTION_ARRAY(quest).phone_checkbox = unchecked Then housing_info_txt = housing_info_txt & "   [ ] Phone "
					If FORM_QUESTION_ARRAY(quest).phone_checkbox = checked Then housing_info_txt = housing_info_txt & "   [X] Phone "

					subsidy_info_tx = "Subsidy: " & FORM_QUESTION_ARRAY(quest).subsidy_yn & "    Subsidy Amount: $ " & FORM_QUESTION_ARRAY(quest).subsidy_amount

					Text 30, y_pos, 450, 10, housing_info_txt
					Text 30, y_pos+10, 450, 10, subsidy_info_tx
					y_pos = y_pos + 20
				End If

				'Display information in columns to reduce risk of details extending past dialog
				x_pos = 30
				yes_answer_count = 0
				If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then
					Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
					y_pos = y_pos + 10
					If FORM_QUESTION_ARRAY(quest).answer_is_array = True Then
						For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
							If FORM_QUESTION_ARRAY(quest).item_ans_list(each_caf_unea) = "Yes" Then
								Text x_pos, y_pos, 140, 10, FORM_QUESTION_ARRAY(quest).item_info_list(each_caf_unea) & " - Yes"
								yes_answer_count = yes_answer_count + 1
								x_pos = x_pos + 140
								If x_pos > 310 Then
									x_pos = 30
									y_pos = y_pos + 10
								End If
							End If
						Next
						If yes_answer_count = 1 Or yes_answer_count = 2 Or yes_answer_count = 4 OR yes_answer_count = 5 then y_pos = y_pos + 10
					End If

					If FORM_QUESTION_ARRAY(quest).write_in_info <> "" Then
						Text 30, y_pos, 450, 10, "Write-In: " & FORM_QUESTION_ARRAY(quest).write_in_info
						y_pos = y_pos + 10
					End If
					If FORM_QUESTION_ARRAY(quest).interview_notes <> "" Then
						Text 30, y_pos, 450, 10, "Interview Notes: " & FORM_QUESTION_ARRAY(quest).interview_notes
						y_pos = y_pos + 10
					End If
				End If

				'Adjust the utilities information to condense information
				x_pos = 30
				If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then
					first_array = True
					If FORM_QUESTION_ARRAY(quest).answer_is_array = True Then
						For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
							If FORM_QUESTION_ARRAY(quest).item_ans_list(each_caf_unea) = "Yes" Then
								If first_array = True Then
									Text 25, y_pos, 450, 10, FORM_QUESTION_ARRAY(quest).number & "." & FORM_QUESTION_ARRAY(quest).dialog_phrasing
									y_pos = y_pos + 10
									first_array = False
								End If
								Text x_pos, y_pos, 105, 10, FORM_QUESTION_ARRAY(quest).item_info_list(each_caf_unea) & " - Yes"
								x_pos = x_pos + 115
								If x_pos > 260 Then
									x_pos = 30
									y_pos = y_pos + 10
								End If
							End If
						Next
					End If
				End If
			Next
			y_pos = y_pos + 10

			Text 15, y_pos, 305, 10, "Add notes or other details in making the expedited determination:"
			EditBox 15, y_pos+10, 450, 15, exp_det_notes
		ElseIf page_display = emergency_questions Then

			Text 505, emer_pos, 60, 10, "EMER Q"

			GroupBox 5, 5, 475, 115, "Emergency Details"
			Text 15, 20, 145, 10, "Is the resident experiencing an emergency?"
			DropListBox 190, 20, 30, 15, " "+chr(9)+"Yes"+chr(9)+"No", resident_emergency_yn
			Text 15, 40, 160, 10, "What emergency is the resident is experiencing?"
			ComboBox 190, 40, 210, 25, "Select or Type"+chr(9)+"Eviction"+chr(9)+"Forced Move"+chr(9)+"Foreclosure"+chr(9)+"Utility Disconnect"+chr(9)+"Home Repairs"+chr(9)+"Property Taxes"+chr(9)+"Bus Ticket"+chr(9)+emergency_type, emergency_type
			Text 15, 65, 130, 10, "Discussion of emergency with resident:"
			EditBox 190, 60, 210, 15, emergency_discussion
			Text 15, 85, 170, 10, "What amount is needed to resolve the emergency?"
			EditBox 190, 80, 45, 15, emergency_amount
			Text 15, 105, 170, 10, "What is the deadline to resolve the emergency?"
			EditBox 190, 100, 45, 15, emergency_deadline
		' ElseIf page_display =  Then
		ElseIf page_display = show_pg_last Then
			If second_signature_detail = "Select or Type" or second_signature_detail = "" Then
				If UBound(HH_MEMB_ARRAY, 2) = 0 Then second_signature_detail = "Not Required"
			End If
			Text 498, last_pos, 60, 10, "CAF Last Page"

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
		    ComboBox 250, 80, 200, 45, all_the_clients+chr(9)+signature_person, signature_person
		    Text 10, 105, 90, 10, "Signature of Other Adult"
		    ComboBox 105, 100, 110, 45, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Not Required"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+second_signature_detail, second_signature_detail
		    Text 220, 105, 25, 10, "person"
		    ComboBox 250, 100, 200, 45, all_the_clients+chr(9)+second_signature_person, second_signature_person

			Text 10, 125, 320, 20, "Only select 'Accepted Verbally' if you are the one accepting the signature verbally. For signatures accepted by another worker, indicate the signature as completed."
			Text 335, 125, 50, 10, "Interview Date:"
			EditBox 390, 120, 60, 15, interview_date

			GroupBox 5, 150, 475, 200, "Benefit Detail"
			y_pos = 165
			If interview_questions_clear = False Then
				Text 15, 165, 450, 10, "ADDITIONAL QUESTIONS BEFORE ASSESMENT IS COMPLETE."
				y_pos = 185
			End If

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

		ElseIf page_display >= 4 or page_display <= last_page_of_questions Then		'This has to be at the end of the ifs because
			' display_count = 1
			y_pos = 10

			For quest = 0 to UBound(FORM_QUESTION_ARRAY)
				If FORM_QUESTION_ARRAY(quest).dialog_page_numb = page_display Then
					' If FORM_QUESTION_ARRAY(quest).dialog_order = display_count Then
					If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), "", True)
					If FORM_QUESTION_ARRAY(quest).answer_is_array = true  Then
						If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), "", True)
						If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_HOUSING_ARRAY, True)
						If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_UTILITIES_ARRAY, True)
						If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_ASSETS_ARRAY, True)
						If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_MSA_ARRAY, True)
						If FORM_QUESTION_ARRAY(quest).info_type = "stwk" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_STWK_ARRAY, True)
					End If
				End If
			Next
			If page_display = 4 and last_page_of_questions => 4 Then Text 505, 67, 55, 10, pg_4_label
			If page_display = 5 and last_page_of_questions => 5 Then Text 505, 82, 55, 10, pg_5_label
			If page_display = 6 and last_page_of_questions => 6 Then Text 505, 97, 55, 10, pg_6_label
			If page_display = 7 and last_page_of_questions => 7 Then Text 505, 112, 55, 10, pg_7_label
			If page_display = 8 and last_page_of_questions => 8 Then Text 505, 127, 55, 10, pg_8_label
			If page_display = 9 and last_page_of_questions => 9 Then Text 505, 142, 55, 10, pg_9_label
			If page_display = 10 and last_page_of_questions => 10 Then Text 505, 157, 55, 10, pg_10_label
			If page_display = 11 and last_page_of_questions => 11 Then Text 505, 172, 55, 10, pg_11_label

		End If


		If page_display <> show_cover_letter and CAF_form = "MNbenefits" Then PushButton 485, 5, 65, 13, "COVER LETTER", cover_letter_btn
		If page_display <> show_pg_one_memb01_and_exp 	Then PushButton 495, 20, 55, 13, "INTVW / CAF 1", caf_page_one_btn
		If page_display <> show_pg_one_address 			Then PushButton 495, 35, 55, 13, "CAF ADDR", caf_addr_btn
		If page_display <> show_pg_memb_list 			Then PushButton 495, 50, 55, 13, "CAF MEMBs", caf_membs_btn
		btn_pos = 65
		If page_display <> 4 									Then PushButton 495, btn_pos, 		55, 13, pg_4_label, caf_q_pg_4_btn
		If page_display <> 5 and last_page_of_questions => 5 	Then PushButton 495, btn_pos + 15, 	55, 13, pg_5_label, caf_q_pg_5_btn
		If page_display <> 6 and last_page_of_questions => 6 	Then PushButton 495, btn_pos + 30, 	55, 13, pg_6_label, caf_q_pg_6_btn
		If page_display <> 7 and last_page_of_questions => 7 	Then PushButton 495, btn_pos + 45, 	55, 13, pg_7_label, caf_q_pg_7_btn
		If page_display <> 8 and last_page_of_questions => 8 	Then PushButton 495, btn_pos + 60, 	55, 13, pg_8_label, caf_q_pg_8_btn
		If page_display <> 9 and last_page_of_questions => 9 	Then PushButton 495, btn_pos + 75, 	55, 13, pg_9_label, caf_q_pg_9_btn
		If page_display <> 10 and last_page_of_questions => 10 	Then PushButton 495, btn_pos + 90, 	55, 13, pg_10_label, caf_q_pg_10_btn
		If page_display <> 11 and last_page_of_questions => 11 	Then PushButton 495, btn_pos + 105, 55, 13, pg_11_label, caf_q_pg_11_btn
		btn_pos = (last_page_of_questions * 15) + 20

		If page_display <> show_qual 					Then PushButton 495, btn_pos, 		55, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> emergency_questions			Then PushButton 495, btn_pos + 15, 	55, 13, "EMER Q", emer_questions_btn
		btn_pos = btn_pos + 30

		If discrepancies_exist = True Then
			If page_display <> discrepancy_questions 	Then PushButton 495, btn_pos, 55, 13, "Clarifications", discrepancy_questions_btn
			btn_pos = btn_pos + 15
		End If
		If expedited_determination_needed = True Then
			If page_display <> expedited_determination Then PushButton 495, btn_pos, 55, 13, "EXPEDITED", expedited_determination_btn
			btn_pos = btn_pos + 15
		End If

		If page_display <> show_pg_last 				Then PushButton 495, btn_pos, 55, 13, "CAF Last Page", caf_last_page_btn
		PushButton 485, btn_pos+25, 66, 13, "Prog Request", program_requests_btn


		PushButton 10, 365, 130, 15, "Interview Ended - INCOMPLETE", incomplete_interview_btn
		If run_by_interview_team = False Then PushButton 140, 365, 130, 15, "View Verifications", verif_button
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
			If page_display = show_pg_memb_list Then
                selected_memb = i
                update_pers = True
            End If
		End If
        If ButtonPressed = HH_MEMB_ARRAY(button_two, i) Then
            HH_MEMB_ARRAY(ignore_person, i) = True
            selected_memb = 0
        End If
	Next


	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If ButtonPressed = FORM_QUESTION_ARRAY(quest).verif_btn Then
			call FORM_QUESTION_ARRAY(quest).capture_verif_detail()
		End If
		If ButtonPressed = FORM_QUESTION_ARRAY(quest).prefil_btn Then
			For i = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_ans_list)
				FORM_QUESTION_ARRAY(quest).item_ans_list(i) = "No"
			Next
			If FORM_QUESTION_ARRAY(quest).sub_phrase <> "" Then FORM_QUESTION_ARRAY(quest).sub_answer = "No"
		End If

		If ButtonPressed = FORM_QUESTION_ARRAY(quest).add_to_array_btn Then
			another_job = ""
			tally = 0
			for each_item = 0 to UBOUND(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				tally = tally + 1
				blank_item = true
				If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then
					If FORM_QUESTION_ARRAY(quest).detail_business(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_monthly_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_value(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_explain(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "expense" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_current(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "winnings" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "changes" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_explain(each_item) <> "" Then blank_item = false
				End If
				If blank_item = true Then another_job = each_item
			Next
			If another_job = "" Then
				another_job = tally
				' MsgBox "another_job - " & another_job & vbCr & IsArray(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				FORM_QUESTION_ARRAY(quest).add_detail_item(another_job)
				FORM_QUESTION_ARRAY(quest).detail_edit_btn(another_job) = 2000 + quest*10 + another_job

			End If
			Call array_details_dlg(quest, another_job)
		End If
		If IsArray(FORM_QUESTION_ARRAY(quest).detail_interview_notes) = true Then
			for each_item = 0 to UBOUND(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				If ButtonPressed = FORM_QUESTION_ARRAY(quest).detail_edit_btn(each_item) Then Call array_details_dlg(quest, each_item)
			next
		End If
	Next


	If ButtonPressed = program_requests_btn Then Call verbal_requests
	If ButtonPressed = open_hsr_manual_transfer_page_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Case-Transfers.aspx"

	If arep_name <> "" Then arep_exists = True
	If ButtonPressed = update_information_btn Then
		If page_display = show_pg_one_address Then
			update_addr = TRUE
			need_to_update_addr = TRUE
		End If
		' If page_display = show_pg_memb_list Then update_pers = TRUE
		If page_display = show_pg_last Then page_display = show_arep_page
		' MsgBox update_arep & " - in dlg move"
	End If
	If ButtonPressed = save_information_btn Then
		If page_display = show_pg_one_address Then update_addr = FALSE
		If page_display = show_pg_memb_list Then update_pers = FALSE
		If page_display = show_arep_page Then page_display = show_pg_last

	End If
	If ButtonPressed = clear_mail_addr_btn Then
		mail_addr_street_full = ""
		mail_addr_city = ""
		mail_addr_state = ""
		mail_addr_zip = ""
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
        pers_next = False                           'Used to navigate the dialog to a MEMBER Update page if an update is required for the member.
		If page_display = show_pg_memb_list Then
            If NOT IsNumeric(selected_memb) Then selected_memb = 0
            For each_memb = selected_memb to UBound(HH_MEMB_ARRAY, 2)
                If HH_MEMB_ARRAY(requires_update, each_memb) Then
                    selected_memb = each_memb
                    update_pers = True
                    pers_next = True
                    Exit For
                End If
            Next
        End If

        If NOT pers_next Then
            page_display = page_display + 1
            If page_display > show_pg_last Then
                page_display = show_pg_last
                ButtonPressed = finish_interview_btn
            End If
            If page_display = discrepancy_questions and discrepancies_exist = False Then page_display = page_display + 1
            If page_display = expedited_determination and expedited_determination_needed = False Then page_display = page_display + 1
        End If
	End If

    If ButtonPressed = cover_letter_btn Then
        page_display = show_cover_letter
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
	If ButtonPressed = caf_q_pg_4_btn Then page_display = 4
	If ButtonPressed = caf_q_pg_5_btn Then page_display = 5
	If ButtonPressed = caf_q_pg_6_btn Then page_display = 6
	If ButtonPressed = caf_q_pg_7_btn Then page_display = 7
	If ButtonPressed = caf_q_pg_8_btn Then page_display = 8
	If ButtonPressed = caf_q_pg_9_btn Then page_display = 9
	If ButtonPressed = caf_q_pg_10_btn Then page_display = 10
	If ButtonPressed = caf_q_pg_11_btn Then page_display = 11

	If ButtonPressed = caf_qual_q_btn Then
		page_display = show_qual
	End If
	If ButtonPressed = caf_last_page_btn Then
		page_display = show_pg_last
	End If
	If ButtonPressed = discrepancy_questions_btn Then
		page_display = discrepancy_questions
	End If
	If ButtonPressed = expedited_determination_btn OR page_display = expedited_determination Then
		If run_by_interview_team = True Then page_display = expedited_determination
		If run_by_interview_team = False Then
			STATS_manualtime = STATS_manualtime + 150
			call display_expedited_dialog
		End If
	End If
	If ButtonPressed = emer_questions_btn Then
		page_display = emergency_questions
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
    If page_display <> show_pg_memb_list or update_pers = False Then
        selected_memb = ""
        update_pers = False
    End If
end function

function display_errors(the_err_msg, execute_nav, show_err_msg_during_movement)
	' MsgBox "the_err_msg" & vbCr & the_err_msg
    If the_err_msg <> "" Then       'If the error message is blank - there is nothing to show.
		If left(the_err_msg, 3) = "~!~" Then the_err_msg = right(the_err_msg, len(the_err_msg) - 3)     'Trimming the message so we don't have a blank array item
        err_array = split(the_err_msg, "~!~")           'making the list of errors an array.

		end_interview = False
		If ButtonPressed = finish_interview_btn Then end_interview = True
		If page_display = show_pg_last AND (ButtonPressed = next_btn OR ButtonPressed = -1) Then end_interview = True

        error_message = ""                              'blanking out variables
        msg_header = ""
        for each message in err_array                   'going through each error message to order them and add headers'

			If show_err_msg_during_movement = False OR end_interview = True Then
	            current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing = trim(current_listing)
	            If current_listing <> msg_header Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
	                If current_listing = "1"  Then tagline = ": Interview"        'Adding a specific tagline to the header for the errors
	                If current_listing = "2"  Then tagline = ": CAF ADDR"
	                If current_listing = "3"  Then tagline = ": CAF MEMBs"

					If current_listing = "4"  									Then tagline = ": " & pg_4_label
					If last_page_of_questions => 5 and current_listing = "5" 	Then tagline = ": " & pg_5_label
					If last_page_of_questions => 6 and current_listing = "6" 	Then tagline = ": " & pg_6_label
					If last_page_of_questions => 7 and current_listing = "7" 	Then tagline = ": " & pg_7_label
					If last_page_of_questions => 8 and current_listing = "8" 	Then tagline = ": " & pg_8_label
					If last_page_of_questions => 9 and current_listing = "9" 	Then tagline = ": " & pg_9_label
					If last_page_of_questions => 10 and current_listing = "10" 	Then tagline = ": " & pg_10_label
					If last_page_of_questions => 11 and current_listing = "11" 	Then tagline = ": " & pg_11_label

					If current_listing = qual_numb 		Then tagline = ": CAF QUAL Q"
					If current_listing = emer_numb 		Then tagline = ": Emergency Questions"
					If current_listing = discrep_num 	Then tagline = ": Clarifications"
					If current_listing = last_num 		Then tagline = ": CAF Last Page"

					error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
	            End If
	            if msg_header = "" Then back_to_dialog = current_listing
	            msg_header = current_listing        'setting for the next loop

	            message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

	            error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
			ElseIf show_err_msg_during_movement = TRUE Then
				page_to_review = page_display & ""
				If page_display = show_qual Then page_to_review = qual_numb
				If page_display = emergency_questions Then page_to_review = emer_numb
				If page_display = discrepancy_questions Then page_to_review = discrep_num
				If page_display = expedited_determination Then page_to_review = exp_num
				If page_display = show_pg_last Then page_to_review = last_num
				page_to_review = trim(page_to_review) & ""

                If page_to_review = "3" Then                'This will navigate to the first member with an error on the CAF MEMBs page
                    selected_memb = err_selected_memb
                    update_pers = False
                    If IsNumeric(err_selected_memb) Then update_pers = True
                End If

				current_listing = left(message, 2)          'This is the dialog the error came from
				current_listing =  trim(current_listing)

				If current_listing = page_to_review Then                   'this is comparing to the dialog from the last message - if they don't match, we need a new header entered
	                If current_listing = "1"  Then tagline = ": Interview"        'Adding a specific tagline to the header for the errors
	                If current_listing = "2"  Then tagline = ": CAF ADDR"
	                If current_listing = "3"  Then tagline = ": CAF MEMBs"

					If current_listing = "4"  									Then tagline = ": " & pg_4_label
					If last_page_of_questions => 5 and current_listing = "5" 	Then tagline = ": " & pg_5_label
					If last_page_of_questions => 6 and current_listing = "6" 	Then tagline = ": " & pg_6_label
					If last_page_of_questions => 7 and current_listing = "7" 	Then tagline = ": " & pg_7_label
					If last_page_of_questions => 8 and current_listing = "8" 	Then tagline = ": " & pg_8_label
					If last_page_of_questions => 9 and current_listing = "9" 	Then tagline = ": " & pg_9_label
					If last_page_of_questions => 10 and current_listing = "10" 	Then tagline = ": " & pg_10_label
					If last_page_of_questions => 11 and current_listing = "11" 	Then tagline = ": " & pg_11_label

					If current_listing = qual_numb 		Then tagline = ": CAF QUAL Q"
					If current_listing = emer_numb 		Then tagline = ": Emergency Questions"
					If current_listing = discrep_num 	Then tagline = ": Clarifications"
					If current_listing = last_num 		Then tagline = ": CAF Last Page"
					If error_message = "" Then error_message = error_message & vbNewLine & vbNewLine & "----- Dialog " & current_listing & tagline & " -------"    'This is the header verbiage being added to the message text.
					message = replace(message, "##~##", vbCR)       'This is notation used in the creation of the message to indicate where we want to have a new line.'

					error_message = error_message & vbNewLine & right(message, len(message) - 3)        'Adding the error information to the message list.
					' MsgBox "error_message - " & error_message & vbCr & "page_to_review - " & page_to_review & vbCr & "current_listing - " & current_listing
				End If
			End If
        Next
		If error_message = "" then the_err_msg = ""

        'This is the display of all of the messages.
		show_msg = False
        If show_err_msg_during_movement = True Then show_msg = True
		If page_display = show_pg_last Then
			show_msg = False
			If ButtonPressed = finish_interview_btn OR ButtonPressed = next_btn OR ButtonPressed = -1 Then show_msg = True
		End If

		If page_display = discrepancy_questions Then show_msg = False
		If ButtonPressed = exp_income_guidance_btn Then show_msg = False
		If ButtonPressed = incomplete_interview_btn Then show_msg = False
		If ButtonPressed = verif_button Then show_msg = False
		If ButtonPressed = open_hsr_manual_transfer_page_btn Then show_msg = False
		If ButtonPressed >= 500 AND ButtonPressed < 1200 Then show_msg = False
		If ButtonPressed >= 4000 Then show_msg = False

		If page_display = emergency_questions and (ButtonPressed = next_btn OR ButtonPressed = -1) Then show_msg = True
		If error_message = "" Then show_msg = False

		If ButtonPressed = finish_interview_btn Then show_msg = True


		If show_msg = True Then view_errors = MsgBox("In order to complete the script and CASE/NOTE, additional details need to be added or refined. Please review and update." & vbNewLine & error_message, vbCritical, "Review detail required in Dialogs")
		If show_msg = False then the_err_msg = ""

        'The function can be operated without moving to a different dialog or not. The only time this will be activated is at the end of dialog 8.
        If execute_nav = TRUE AND show_err_msg_during_movement = False Then
            If IsNumeric(back_to_dialog) = true Then
				'This calls another function to go to the first dialog that had an error
				back_to_dialog = back_to_dialog * 1
				If back_to_dialog >= 4 and back_to_dialog <= last_page_of_questions Then
					If back_to_dialog = 4 Then ButtonPressed = caf_q_pg_4_btn
					If back_to_dialog = 5 Then ButtonPressed =  caf_q_pg_5_btn
					If back_to_dialog = 6 Then ButtonPressed =  caf_q_pg_6_btn
					If back_to_dialog = 7 Then ButtonPressed =  caf_q_pg_7_btn
					If back_to_dialog = 8 Then ButtonPressed =  caf_q_pg_8_btn
					If back_to_dialog = 9 Then ButtonPressed =  caf_q_pg_9_btn
					If back_to_dialog = 10 Then ButtonPressed =  caf_q_pg_10_btn
					If back_to_dialog = 11 Then ButtonPressed =  caf_q_pg_11_btn
				End If
				If back_to_dialog = show_pg_one_memb01_and_exp	Then ButtonPressed = caf_page_one_btn
				If back_to_dialog = show_pg_one_address			Then ButtonPressed = caf_addr_btn
				If back_to_dialog = show_pg_memb_list			Then ButtonPressed = caf_membs_btn
				If back_to_dialog = show_qual					Then ButtonPressed = caf_qual_q_btn
				If back_to_dialog = show_pg_last				Then ButtonPressed = caf_last_page_btn
				If back_to_dialog = discrepancy_questions		Then ButtonPressed = discrepancy_questions_btn
				If back_to_dialog = expedited_determination		Then ButtonPressed = expedited_determination_btn
				If back_to_dialog = emergency_questions 		Then ButtonPressed = emer_questions_btn

				Call dialog_movement          'this is where the navigation happens
			End If
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

	If first_time_in_exp_det = True Then
		heat_expense = False
		ac_expense = False
		electric_expense = False
		phone_expense = False

		For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
			If FORM_QUESTION_ARRAY(each_question).info_type = "jobs" or FORM_QUESTION_ARRAY(each_question).detail_source = "jobs" Then
				If FORM_QUESTION_ARRAY(each_question).caf_answer = "Yes" Then jobs_income_yn = "Yes"
				If FORM_QUESTION_ARRAY(each_question).caf_answer = "No" Then jobs_income_yn = "No"
				If FORM_QUESTION_ARRAY(each_question).detail_array_exists = True Then
					exp_job_count = 0
					For each_caf_job = 0 to UBound(FORM_QUESTION_ARRAY(each_question).detail_business)
						If FORM_QUESTION_ARRAY(each_question).detail_business(each_caf_job) <> "" Then
							ReDim Preserve EXP_JOBS_ARRAY(jobs_notes_const, exp_job_count)
							EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count) = FORM_QUESTION_ARRAY(each_question).detail_resident_name(each_caf_job)
							If len(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count)) > 5 Then EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count) = right(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count), len(EXP_JOBS_ARRAY(jobs_employee_const, exp_job_count))-5)

							EXP_JOBS_ARRAY(jobs_employer_const, exp_job_count) = FORM_QUESTION_ARRAY(each_question).detail_business(each_caf_job)
							EXP_JOBS_ARRAY(jobs_wage_const, exp_job_count) = FORM_QUESTION_ARRAY(each_question).detail_hourly_wage(each_caf_job)

							If IsNumeric(FORM_QUESTION_ARRAY(each_question).detail_monthly_amount(each_caf_job)) = True and IsNumeric(FORM_QUESTION_ARRAY(each_question).detail_hourly_wage(each_caf_job)) = True Then
								If FORM_QUESTION_ARRAY(each_question).detail_hourly_wage(each_caf_job) > 0 Then      'making sure we are not dividing by zero. I will not be defaulting to a zero income job - no autofils
									monthly_hours = FORM_QUESTION_ARRAY(each_question).detail_monthly_amount(each_caf_job)/FORM_QUESTION_ARRAY(each_question).detail_hourly_wage(each_caf_job)
									weekly_hours = monthly_hours/4
									EXP_JOBS_ARRAY(jobs_hours_const, exp_job_count) = weekly_hours
									EXP_JOBS_ARRAY(jobs_frequency_const, exp_job_count) = "Weekly"
								End If
							End If

							exp_job_count = exp_job_count + 1
						End If
					Next
				End If
			End If

			If InStr(FORM_QUESTION_ARRAY(each_question).note_phrasing, "Is anyone self-employed?") <> 0 Then
				If FORM_QUESTION_ARRAY(each_question).caf_answer = "Yes" Then busi_income_yn = "Yes"
				If FORM_QUESTION_ARRAY(each_question).caf_answer = "No" Then busi_income_yn = "No"
			End If

			If FORM_QUESTION_ARRAY(each_question).info_type = "unea" Then
				unea_income_yn = "No"
				exp_unea_count = 0
				If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
					For each_unea = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
						If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_unea) = "Yes" Then
							ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
							EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = FORM_QUESTION_ARRAY(each_question).item_info_list(each_unea)
							EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = FORM_QUESTION_ARRAY(each_question).item_detail_list(each_unea)
							exp_unea_count = exp_unea_count + 1
							unea_income_yn = "Yes"
						End If
					Next
				End If
			End If
			If FORM_QUESTION_ARRAY(each_question).detail_source = "unea" Then
				unea_income_yn = "No"
				exp_unea_count = 0
				If FORM_QUESTION_ARRAY(each_question).detail_array_exists = True Then
					each_caf_unea = 0
					For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(each_question).detail_type)
						If FORM_QUESTION_ARRAY(each_question).detail_type(each_caf_unea) <> "" Then
							ReDim Preserve EXP_UNEA_ARRAY(unea_notes_const, exp_unea_count)
							EXP_UNEA_ARRAY(unea_info_const, exp_unea_count) = FORM_QUESTION_ARRAY(each_question).detail_type(each_caf_unea)
							EXP_UNEA_ARRAY(unea_monthly_earnings_const, exp_unea_count) = FORM_QUESTION_ARRAY(each_question).detail_amount(each_caf_unea)
							exp_unea_count = exp_unea_count + 1
							unea_income_yn = "Yes"
						End If
					Next
				End If
			End If


			If FORM_QUESTION_ARRAY(each_question).info_type = "assets" Then
				If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
					For each_acct = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
						If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_acct) = "Cash" Then
							If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_acct) = "Yes" Then cash_amount_yn = "Yes"
							If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_acct) = "No" Then cash_amount_yn = "No"
						End If
						If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_acct) = "Bank Accounts" Then
							If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_acct) = "Yes" Then bank_account_yn = "Yes"
							If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_acct) = "No" Then bank_account_yn = "No"
						End If
					Next
				End If
			End If
			If FORM_QUESTION_ARRAY(each_question).detail_source = "assets" Then
				bank_account_yn = "No"
				If FORM_QUESTION_ARRAY(each_question).detail_array_exists = True Then
					each_caf_asset = 0
					For each_caf_asset = 0 to UBound(FORM_QUESTION_ARRAY(each_question).detail_type)
						If FORM_QUESTION_ARRAY(each_question).detail_type(each_caf_asset) <> "" Then bank_account_yn = "Yes"
					Next
				End If
			End If

			If FORM_QUESTION_ARRAY(each_question).detail_source = "shel-hest" Then
				If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = checked Then
					heat_expense = True
					ac_expense = True
				End If
				If FORM_QUESTION_ARRAY(each_question).electric_checkbox = checked Then electric_expense = True
				If FORM_QUESTION_ARRAY(each_question).phone_checkbox = checked Then phone_expense = True
				determined_shel = FORM_QUESTION_ARRAY(each_question).housing_payment
			End If


			If FORM_QUESTION_ARRAY(each_question).info_type = "housing" Then
				If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
					For each_shel = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
						If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_shel) = "Yes" Then
							If FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel) = "Rent" 								Then rent_amount = exp_q_3_rent_this_month
							If FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel) = "Mortgage/contract for deed payment" 	Then mortgage_amount = exp_q_3_rent_this_month
							If FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel) = "Room and/or Board" 					Then room_amount = exp_q_3_rent_this_month
							If FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel) = "Homeowner's insurance" 				Then insurance_amount = exp_q_3_rent_this_month
							If FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel) = "Real estate taxes" 					Then tax_amount = exp_q_3_rent_this_month
						End If
					Next
				End If
			End If


			If FORM_QUESTION_ARRAY(each_question).info_type = "utilities" Then
				If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
					For each_util = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
						If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_util) = "Yes" Then
							If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Heat/AC" Then
								heat_expense = True
								ac_expense = True
							End If
							If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Electric" 	Then electric_expense = True
							If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone" 		Then phone_expense = True
						End If
					Next
				End If
			End If
		Next

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

		Call app_month_income_detail(determined_income, income_review_completed, jobs_income_yn, busi_income_yn, unea_income_yn, EXP_JOBS_ARRAY, EXP_BUSI_ARRAY, EXP_UNEA_ARRAY)
		Call app_month_asset_detail(determined_assets, assets_review_completed, cash_amount_yn, bank_account_yn, cash_amount, EXP_ACCT_ARRAY)
		Call app_month_housing_detail(determined_shel, shel_review_completed, rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, subsidy_amount)
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

				If expedited_screening_on_form = True Then
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
				Else
					Text 10, 20, 350, 10, "Expedited Screening Questions not present on this form. No information to Display."
				End If
				' End If
				' If exp_screening_note_found = False Then
				' 	Text 10, 20, 350, 10, "CASE:NOTE for Expedited Screening could not be found. No information to Display."
				' 	Text 10, 30, 350, 10, "Review Application for screening answers"
				' End If
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

				'This section will display the details of the notes the worker has entered into the main portion of the interview script.
				'These details are intended to support update of Expedited Determination information
				y_pos = 215
				GroupBox 5, y_pos, 545, 100, "Interview NOTES entered into the Script already"
				y_pos = y_pos + 15
				' If trim(question_8_interview_notes) <> "" Then
				' 	Text 15, y_pos, 530, 10, "8. Has anyone in the household had a job or been self-employed? " & question_8_interview_notes
				' 	y_pos = y_pos + 10
				' End If

				' If y_pos = 230 Then
				' 	Text 15, y_pos, 530, 10, "No details entered into Interview Notes sections of relevant questions (8, 9, 10, 12, 14, 15, 20) and no specific job details were entered in question 9."
				' 	y_pos = y_pos + 10
				' End If

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

		If ButtonPressed = knowledge_now_support_btn Then
			Call send_support_email_to_KN
			STATS_manualtime = STATS_manualtime + 300
		End If
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

Function display_exemptions() 'A message box showing exemptions from SNAP work rules
	wreg_exemptions = msgbox("Individuals in your household may not have to follow these General Work Rules if [you/they] are:" & vbCr & vbCr &_
				  "* Younger than 16 or older than 59," & vbCr &_
	     		  "* Taking care of a child younger than 6 or someone who needs helps caring for themselves, " & vbCr &_
	     		  "* Already working at least 30 hours a week," & vbCr &_
	     		  "* Already earning $217.50 or more per week," & vbCr &_
	     		  "* Receiving unemployment benefits, or you applied for unemployment benefits," & vbCr &_
	     		  "* Not working because of a physical illness, injury, disability, or surgery recovery," & vbCr &_
	     		  "* Not working due to a mental health illness, disorder, or health condition," & vbCr &_
				  "* Are homeless," & vbCr &_
				  "* A victim of domestic violence," & vbCr &_
				  "* Going to school, college, or a training program at least half time," & vbCr &_
				  "* Meeting the work rules for Minnesota Family Investment Program (MFIP) or DWP (Divisionary Work Program (DWP)," & vbCr &_
				  "* Not working due to a substance use disorder or addiction dependency, or" & vbCr &_
				  "* Participating in a drug or alcohol addiction treatment program." & vbCr & vbCr &_
				  "Press yes if you reviewed exemptions with the resident, press no to return to the previous dialog without review." & vbCr &_
				  "Press 'Cancel' to end the script run.", vbYesNoCancel+ vbQuestion, "Work Rules Reviewed")
		If wreg_exemptions = vbCancel then cancel_confirmation
	If wreg_exemptions = vbYes then work_exemptions_reviewed = true
End Function

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

function program_process_selection()
	dlg_len = 100
	y_pos = 75
	If cash_request = True Then dlg_len = dlg_len + 20
	If snap_request = True Then dlg_len = dlg_len + 20
	If emer_request = True Then dlg_len = dlg_len + 20
	If grh_request = True Then dlg_len = dlg_len + 20
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 210, dlg_len, "CAF Process"
		Text 10, 10, 205, 20, "Interviews are completed on cases when programs are initially requested and at annual renewal for SNAP and MFIP."
		Text 10, 35, 210, 20, "To correctly identify the information needed, each program needs to be associated with an Application or Renewal process."

		Text 10, 60, 35, 10, "Program"
		Text 80, 60, 50, 10, "CAF Process"
		Text 155, 60, 50, 10, "Recert MM/YY"
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
		If grh_request = True Then
			Text 10, y_pos + 5, 20, 10, "GRH"
			DropListBox 80, y_pos, 65, 45, "Select One..."+chr(9)+"Application"+chr(9)+"Renewal", the_process_for_grh
			EditBox 155, y_pos, 20, 15, next_grh_revw_mo
			EditBox 180, y_pos, 20, 15, next_grh_revw_yr
			y_pos = y_pos + 20
		End If
		y_pos = y_pos + 5
		Text 10, y_pos+5, 125, 10, "(The programs do not need to match.)"
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
				If type_of_cash = "?" Then err_msg = err_msg & vbNewLine & "* Indicate if Cash request in for a Family or Adult program."
			End If
			If snap_request = True Then
				If the_process_for_snap = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the SNAP program is at application or renewal."
				If the_process_for_snap = "Renewal" AND (len(next_snap_revw_mo) <> 2 or len(next_snap_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For SNAP at renewal, enter the footer month and year the of the renewal."
			End If
			If emer_request = True Then
				If type_of_emer = "?" Then err_msg = err_msg & vbNewLine & "* Indicate if EMER request in EA or EGA"
			End If
			If grh_request = True Then
				If the_process_for_grh = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select if the GRH program is at application or renewal."
				If the_process_for_grh = "Renewal" AND (len(next_grh_revw_mo) <> 2 or len(next_grh_revw_yr) <> 2) Then err_msg = err_msg & vbNewLine & "* For GRH at renewal, enter the footer month and year the of the renewal."
			End If


			IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** Please resolve to continue ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
	Call check_for_MAXIS(False)

	If snap_request = True AND the_process_for_snap = "Application" Then expedited_determination_needed = True
	If snap_status = "PENDING" Then expedited_determination_needed = True
	If type_of_cash = "Adult" Then family_cash_case_yn = "No"
	If type_of_cash = "Family" Then family_cash_case_yn = "Yes"
	ButtonPressed = return_btn
end function

function set_initial_exp_simplified()
	exp_det_income = 0
	exp_det_assets = 0
	exp_det_housing = 0
	exp_det_utilities = 0
	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If InStr(FORM_QUESTION_ARRAY(quest).dialog_phrasing, "self-employed") <> 0 AND FORM_QUESTION_ARRAY(quest).info_type = "single-detail" Then
			If IsNumeric(FORM_QUESTION_ARRAY(quest).sub_answer) = true Then exp_det_income = exp_det_income + FORM_QUESTION_ARRAY(quest).sub_answer
		End If
		If FORM_QUESTION_ARRAY(quest).detail_array_exists = true Then
			for each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)

				If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then
					If IsNumeric(FORM_QUESTION_ARRAY(quest).detail_monthly_amount(each_item)) = true Then
						exp_det_income = exp_det_income + FORM_QUESTION_ARRAY(quest).detail_monthly_amount(each_item)
					ElseIf IsNumeric(FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item)) = true AND IsNumeric(FORM_QUESTION_ARRAY(quest).detail_hours_per_week(each_item)) = true Then
						exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item)*FORM_QUESTION_ARRAY(quest).detail_hours_per_week(each_item)*4.3)
					End If
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then
					If IsNumeric(FORM_QUESTION_ARRAY(quest).detail_value(each_item)) = True Then exp_det_assets = exp_det_assets + FORM_QUESTION_ARRAY(quest).detail_value(each_item)
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then
					If IsNumeric(FORM_QUESTION_ARRAY(quest).detail_amount(each_item)) = true Then
						If FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) = "Weekly" Then
							exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_amount(each_item)*4.3)
						ElseIf FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) = "Bi-weekly" Then
							exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_amount(each_item)*2.15)
						ElseIf FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) = "Semi-monthly" Then
							exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_amount(each_item)*2)
						ElseIf FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) = "Once a month" Then
							exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_amount(each_item))
						Else
							exp_det_income = exp_det_income + (FORM_QUESTION_ARRAY(quest).detail_amount(each_item))
						End If
					End If
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then
					If IsNumeric(FORM_QUESTION_ARRAY(quest).housing_payment) = True Then exp_det_housing = FORM_QUESTION_ARRAY(quest).housing_payment
					heat_exp_checkbox = FORM_QUESTION_ARRAY(quest).heat_air_checkbox
					ac_exp_checkbox = FORM_QUESTION_ARRAY(quest).heat_air_checkbox
					electric_exp_checkbox = FORM_QUESTION_ARRAY(quest).electric_checkbox
					phone_exp_checkbox = FORM_QUESTION_ARRAY(quest).phone_checkbox
				End If
			next
		End If
		If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then
			If FORM_QUESTION_ARRAY(quest).answer_is_array = True Then
				For each_caf_unea = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
					If FORM_QUESTION_ARRAY(quest).item_info_list(each_caf_unea) <> "" and FORM_QUESTION_ARRAY(quest).item_ans_list(each_caf_unea) = "Yes" Then
						If IsNumeric(FORM_QUESTION_ARRAY(quest).item_detail_list(each_caf_unea)) = true Then exp_det_income = exp_det_income + FORM_QUESTION_ARRAY(quest).item_detail_list(each_caf_unea)
					End If
				Next
			End If
		End If
		If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then
			If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
				For each_shel = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
					If FORM_QUESTION_ARRAY(quest).item_ans_list(each_shel) = "Yes" Then
						If exp_det_housing = 0 and IsNumeric(exp_q_3_rent_this_month) = True Then exp_det_housing = exp_q_3_rent_this_month
					End If
				Next
			End If
		End If

		If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then
			If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
				For each_util = 0 to UBound(FORM_QUESTION_ARRAY(quest).item_info_list)
					If FORM_QUESTION_ARRAY(quest).item_ans_list(each_util) = "Yes" Then
						If FORM_QUESTION_ARRAY(quest).item_note_info_list(each_util) = "Heat/AC" Then
							heat_exp_checkbox = checked
							ac_exp_checkbox = checked
						End If
						If FORM_QUESTION_ARRAY(quest).item_note_info_list(each_util) = "Electric" 	Then electric_exp_checkbox = checked
						If FORM_QUESTION_ARRAY(quest).item_note_info_list(each_util) = "Phone" 		Then phone_exp_checkbox = checked
					End If
				Next
			End If
		End If
	Next
	exp_det_income = FormatNumber(exp_det_income, 2, -1, 0, -1) & ""
	If exp_det_income = "0.00" Then exp_det_income = ""
	exp_det_assets = FormatNumber(exp_det_assets, 2, -1, 0, -1) & ""
	If exp_det_assets = "0.00" Then exp_det_assets = ""
	exp_det_housing = FormatNumber(exp_det_housing, 2, -1, 0, -1) & ""
	If exp_det_housing = "0.00" Then exp_det_housing = ""

	If heat_exp_checkbox = checked OR ac_exp_checkbox = checked Then
		exp_det_utilities = exp_det_utilities + heat_AC_amt
	Else
		If electric_exp_checkbox = checked Then exp_det_utilities = exp_det_utilities + electric_amt
		If phone_exp_checkbox = checked Then exp_det_utilities = exp_det_utilities + phone_amt
	End If
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

function verbal_requests()
	program_marked_on_CAF = False
	all_programs_marked_on_CAF = True
	If CASH_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If SNAP_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If EMER_on_CAF_checkbox = checked Then program_marked_on_CAF = True
	If GRH_on_CAF_checkbox = checked Then program_marked_on_CAF = True

	If CASH_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If SNAP_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If EMER_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False
	If GRH_on_CAF_checkbox = unchecked Then all_programs_marked_on_CAF = False

	dlg_len = 230
	If snap_closed_in_past_30_days = True or snap_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
	If grh_closed_in_past_30_days = True or grh_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
	If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
	If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
	If issued_date <> "" Then dlg_len = dlg_len + 10
	If dlg_len = 230 Then dlg_len = dlg_len + 10

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 316, dlg_len, "Programs to Interview For"
		GroupBox 10, 10, 295, 60, "Form Details:"
		Text 20, 25, 155, 10, CAF_form_name
		Text 20, 40, 125, 10, "CAF Date: " & CAF_datestamp
		Text 190, 15, 95, 10, "Programs Marked on Form:"
		prog_y_pos = 25
		If CASH_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 50, 10, "- CASH"
			prog_y_pos = prog_y_pos + 10
		End If
		If SNAP_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 50, 10, "- SNAP"
			prog_y_pos = prog_y_pos + 10
		End If
		If EMER_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 70, 10, "- EMERGENCY"
			prog_y_pos = prog_y_pos + 10
		End If
		If GRH_on_CAF_checkbox = checked Then
			Text 195, prog_y_pos, 100, 10, "- HOUSING SUPPORT / GRH"
			prog_y_pos = prog_y_pos + 10
		End If
		If prog_y_pos = 25 Then Text 195, prog_y_pos, 100, 10, "NONE"

		If all_programs_marked_on_CAF = False Then
			req_y_pos = 100
			If CASH_on_CAF_checkbox = unchecked Then
				Text 65, req_y_pos, 25, 10, " Cash:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", cash_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If SNAP_on_CAF_checkbox = unchecked Then
				Text 65, req_y_pos, 25, 10, "SNAP:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", snap_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If EMER_on_CAF_checkbox = unchecked Then
				Text 40, req_y_pos, 50, 10, "EMERGENCY:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", emer_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			If GRH_on_CAF_checkbox = unchecked Then
				Text 15, req_y_pos, 75, 10, " HOUSING SUPPORT:"
				DropListBox 90, req_y_pos-5, 60, 45, "No"+chr(9)+"Yes", grh_verbal_request
				req_y_pos = req_y_pos + 15
			End If
			GroupBox 10, 80, 145, req_y_pos-80, "VERBAL PROGRAM REQUESTS "
		End If
		If program_marked_on_CAF = True Then
			wthdrw_y_pos = 100
			If CASH_on_CAF_checkbox = checked Then
				Text 215, wthdrw_y_pos, 25, 10, " Cash:"
				DropListBox 240, wthdrw_y_pos-5, 60, 45, "No"+chr(9)+"Yes", cash_verbal_withdraw
				wthdrw_y_pos = wthdrw_y_pos + 15
			End If
			If SNAP_on_CAF_checkbox = checked Then
				Text 215, wthdrw_y_pos, 25, 10, "SNAP:"
				DropListBox 240, wthdrw_y_pos-5, 60, 45, "No"+chr(9)+"Yes", snap_verbal_withdraw
				wthdrw_y_pos = wthdrw_y_pos + 15
			End If
			If EMER_on_CAF_checkbox = checked Then
				Text 190, wthdrw_y_pos, 50, 10, "EMERGENCY:"
				DropListBox 240, wthdrw_y_pos-5, 60, 45, "No"+chr(9)+"Yes", emer_verbal_withdraw
				wthdrw_y_pos = wthdrw_y_pos + 15
			End If
			If GRH_on_CAF_checkbox = checked Then
				Text 165, wthdrw_y_pos, 75, 10, " HOUSING SUPPORT:"
				DropListBox 240, wthdrw_y_pos-5, 60, 45, "No"+chr(9)+"Yes", grh_verbal_withdraw
				wthdrw_y_pos = wthdrw_y_pos + 15
			End If
			GroupBox 160, 80, 145, wthdrw_y_pos-80, "VERBAL PROGRAM WITHDRAWALS"
		End If
		y_pos = 170
		orig_y_pos = y_pos
		If snap_closed_in_past_30_days = True or snap_closed_in_past_4_months = True Then
			Text 20, y_pos, 260, 10, "SNAP recently closed on " & FS_date_closed & " - " & FS_reason_closed
			y_pos = y_pos + 10
		End If
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
			Text 20, y_pos, 260, 10, cash1_recently_closed_program & " recently closed on " & cash1_date_closed & " - " & cash1_closed_reason
			y_pos = y_pos + 10
		End If
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
			Text 20, y_pos, 260, 10, cash2_recently_closed_program & " recently closed on " & cash2_date_closed & " - " & cash2_closed_reason
			y_pos = y_pos + 10
		End If
		If grh_closed_in_past_30_days = True or grh_closed_in_past_4_months = True Then
			Text 20, y_pos, 285, 10, "GRH/HS recently closed on " & GRH_date_closed & " - " & GRH_reason_closed
			y_pos = y_pos + 10
		End If
		If issued_date <> "" Then
			Text 20, y_pos, 260, 10, "EMER last issued on " & issued_date & " (" & issued_prog & ")"
			y_pos = y_pos + 10
		End If
		If y_pos = orig_y_pos Then
			Text 20, y_pos, 260, 10, "NO RECENT PROGRAM HISTORY TO NOTE"
			y_pos = y_pos + 10
		End If
		GroupBox 10, orig_y_pos-10, 295, y_pos-orig_y_pos+10, "PROGRAM HISTORY"
		Text 10, y_pos+5, 220, 10, "Additional Notes about Verbal Program Requests or Withdrawals"
		EditBox 10, y_pos+15, 295, 15, verbal_request_notes
		ButtonGroup ButtonPressed
			' OkButton 195, 200, 50, 15
			' CancelButton 255, 200, 50, 15
			PushButton 255, y_pos+35, 50, 15, "Return", return_btn
	EndDialog

	Do
		dialog Dialog1
		cancel_confirmation

		verbal_request_notes = trim(verbal_request_notes)

	Loop until ButtonPressed = return_btn
	ButtonPressed = ""

	cash_request = False
	snap_request = False
	emer_request = False
	grh_request = False
	If CASH_on_CAF_checkbox = checked OR cash_verbal_request = "Yes" Then cash_request = True
	If SNAP_on_CAF_checkbox = checked OR snap_verbal_request = "Yes" Then snap_request = True
	If EMER_on_CAF_checkbox = checked OR emer_verbal_request = "Yes" Then emer_request = True
	If GRH_on_CAF_checkbox = checked OR grh_verbal_request = "Yes" Then grh_request = True

	run_process_selection = False
	If cash_request = True Then
		If type_of_cash = "?" or type_of_cash = "" Then run_process_selection = True
		If the_process_for_cash = "Select One..." or the_process_for_cash = "" Then run_process_selection = True
	End If
	If snap_request = True Then
		If the_process_for_snap = "Select One..." or the_process_for_snap = "" Then run_process_selection = True
	End If
	If emer_request = True Then
		If type_of_emer = "?" or type_of_emer = "" Then run_process_selection = True
		If the_process_for_emer = "Select One..." or the_process_for_emer = "" Then run_process_selection = True
	End If
	If grh_request = True Then
		If the_process_for_grh = "Select One..." or the_process_for_grh = "" Then run_process_selection = True
	End If

	If run_process_selection = True Then call program_process_selection

	ButtonPressed = return_btn
end function

function save_your_work()
'This function records the variables into a txt file so that it can be retrieved by the script if run later.
	Call save_form_details(FORM_QUESTION_ARRAY)

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

			objTextStream.WriteLine "VERB - CASH - " & cash_verbal_request
			objTextStream.WriteLine "VERB - GRHS - " & grh_verbal_request
			objTextStream.WriteLine "VERB - SNAP - " & snap_verbal_request
			objTextStream.WriteLine "VERB - EMER - " & emer_verbal_request
			objTextStream.WriteLine "WTDR - CASH - " & cash_verbal_withdraw
			objTextStream.WriteLine "WTDR - GRHS - " & grh_verbal_withdraw
			objTextStream.WriteLine "WTDR - SNAP - " & snap_verbal_withdraw
			objTextStream.WriteLine "WTDR - EMER - " & emer_verbal_withdraw

			If CASH_on_CAF_checkbox = checked Then objTextStream.WriteLine "CASH PROG CHECKED"
			If GRH_on_CAF_checkbox = checked Then objTextStream.WriteLine "GRHS PROG CHECKED"
			If SNAP_on_CAF_checkbox = checked Then objTextStream.WriteLine "SNAP PROG CHECKED"
			If EMER_on_CAF_checkbox = checked Then objTextStream.WriteLine "EMER PROG CHECKED"

			objTextStream.WriteLine "CASH - TYPE - " & type_of_cash
			objTextStream.WriteLine "PROC - CASH - " & the_process_for_cash
			objTextStream.WriteLine "CASH - RVMO - " & next_cash_revw_mo
			objTextStream.WriteLine "CASH - RVYR - " & next_cash_revw_yr

			objTextStream.WriteLine "PROC - SNAP - " & the_process_for_snap
			objTextStream.WriteLine "SNAP - RVMO - " & next_snap_revw_mo
			objTextStream.WriteLine "SNAP - RVYR - " & next_snap_revw_yr

			objTextStream.WriteLine "PROC - GRHS - " & the_process_for_grh
			objTextStream.WriteLine "GRHS - RVMO - " & next_grh_revw_mo
			objTextStream.WriteLine "GRHS - RVYR - " & next_grh_revw_yr

			objTextStream.WriteLine "EMER - TYPE - " & type_of_emer
			objTextStream.WriteLine "PROC - EMER - " & the_process_for_emer

			objTextStream.WriteLine "PROG - NOTE - " & program_request_notes
			objTextStream.WriteLine "VERB - NOTE - " & verbal_request_notes

			objTextStream.WriteLine "CVR - CMNT - " & additional_application_comments
			objTextStream.WriteLine "CVR - INCM - " & additional_income_comments
			objTextStream.WriteLine "CVR - NOTE - " & cover_letter_interview_notes

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
			objTextStream.WriteLine "ADR - RESI - UPD - " & need_to_update_addr

			objTextStream.WriteLine "ADR - RESI - RES - " & reservation_yn
			objTextStream.WriteLine "ADR - RESI - NAM - " & reservation_name

			objTextStream.WriteLine "ADR - RESI - HML - " & homeless_yn

			objTextStream.WriteLine "ADR - RESI - LIV - " & living_situation

			objTextStream.WriteLine "ADR - HOUS - LIC - " & licensed_facility
			objTextStream.WriteLine "ADR - HOUS - MEA - " & meal_provided
			objTextStream.WriteLine "ADR - HOUS - NAM - " & residence_name_phone

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
			objTextStream.WriteLine "ADR - TEXT - " & send_text
			objTextStream.WriteLine "ADR - EMAL - " & send_email

            objTextStream.WriteLine "MEMB - ALLYN - " & all_members_listed_yn
            objTextStream.WriteLine "MEMB - ALLNT - " & all_members_listed_notes
            objTextStream.WriteLine "MEMB - IMNYN - " & all_members_in_MN_yn
            objTextStream.WriteLine "MEMB - IMNNT - " & all_members_in_MN_notes
            objTextStream.WriteLine "MEMB - PRGYN - " & anyone_pregnant_yn
            objTextStream.WriteLine "MEMB - PRGNT - " & anyone_pregnant_notes
            objTextStream.WriteLine "MEMB - MILYN - " & anyone_served_yn
            objTextStream.WriteLine "MEMB - MILNT - " & anyone_served_notes

			objTextStream.WriteLine "PWE - " & pwe_selection

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
			objTextStream.WriteLine "SIG - 04 - " & second_signature_detail
			objTextStream.WriteLine "SIG - 05 - " & second_signature_person
			objTextStream.WriteLine "SIG - 07 - " & client_signed_verbally_yn
			objTextStream.WriteLine "SIG - 08 - " & interview_date
			objTextStream.WriteLine "SIG - 09 - " & verbal_sig_date
			objTextStream.WriteLine "SIG - 10 - " & verbal_sig_time
			objTextStream.WriteLine "SIG - 11 - " & verbal_sig_phone_number

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

			objTextStream.WriteLine "SIMP EXPDET - 01 - " & exp_det_income
			objTextStream.WriteLine "SIMP EXPDET - 02 - " & exp_det_assets
			objTextStream.WriteLine "SIMP EXPDET - 03 - " & exp_det_housing
			objTextStream.WriteLine "SIMP EXPDET - 04 - " & exp_det_utilities
			objTextStream.WriteLine "SIMP EXPDET - 05 - " & exp_det_notes
			objTextStream.WriteLine "SIMP EXPDET - 06 - " & expedited_viewed

            If verif_snap_checkbox = checked then objTextStream.WriteLine "verif_snap_checkbox"
			If heat_exp_checkbox = checked then objTextStream.WriteLine "heat_exp_checkbox"
			If ac_exp_checkbox = checked then objTextStream.WriteLine "ac_exp_checkbox"
			If electric_exp_checkbox = checked then objTextStream.WriteLine "electric_exp_checkbox"
			If phone_exp_checkbox = checked then objTextStream.WriteLine "phone_exp_checkbox"
			If none_exp_checkbox = checked then objTextStream.WriteLine "none_exp_checkbox"

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

			'Emergency questions
			objTextStream.WriteLine "EMER EXP - " & resident_emergency_yn
			objTextStream.WriteLine "EMER TYPE - " & emergency_type
			objTextStream.WriteLine "EMER DISCUSSION - " & emergency_discussion
			objTextStream.WriteLine "EMER AMOUNT - " & emergency_amount
			objTextStream.WriteLine "EMER DEADLINE - " & emergency_deadline

			'Interview HSR E&T Questions
			objTextStream.WriteLine "SUMM INTNOW - " & interested_in_job_assistance_now
			objTextStream.WriteLine "SUMM NAMENOW - " & interested_names_now
			objTextStream.WriteLine "SUMM INTFUT - " & interested_in_job_assistance_future
			objTextStream.WriteLine "SUMM NAMEFUT - " & interested_names_future

			'R&R
			If DHS_4163_checkbox = checked Then objTextStream.WriteLine "DHS_4163_checkbox"
			If DHS_3315A_checkbox = checked Then objTextStream.WriteLine "DHS_3315A_checkbox"
			If DHS_3979_checkbox = checked Then objTextStream.WriteLine "DHS_3979_checkbox"
			If DHS_2759_checkbox = checked Then objTextStream.WriteLine "DHS_2759_checkbox"
			If DHS_3353_checkbox = checked Then objTextStream.WriteLine "DHS_3353_checkbox"
			If DHS_2920_checkbox = checked Then objTextStream.WriteLine "DHS_2920_checkbox"
			If DHS_3477_checkbox = checked Then objTextStream.WriteLine "DHS_3477_checkbox"
			If DHS_4133_checkbox = checked Then objTextStream.WriteLine "DHS_4133_checkbox"
			If DHS_2647_checkbox = checked Then objTextStream.WriteLine "DHS_2647_checkbox"
			If DHS_2929_checkbox = checked Then objTextStream.WriteLine "DHS_2929_checkbox"
			If DHS_3323_checkbox = checked Then objTextStream.WriteLine "DHS_3323_checkbox"
			If DHS_3393_checkbox = checked Then objTextStream.WriteLine "DHS_3393_checkbox"
			If DHS_3163B_checkbox = checked Then objTextStream.WriteLine "DHS_3163B_checkbox"
			If DHS_2338_checkbox = checked Then objTextStream.WriteLine "DHS_2338_checkbox"
			If DHS_5561_checkbox = checked Then objTextStream.WriteLine "DHS_5561_checkbox"
			If DHS_2961_checkbox = checked Then objTextStream.WriteLine "DHS_2961_checkbox"
			If DHS_2887_checkbox = checked Then objTextStream.WriteLine "DHS_2887_checkbox"
			If DHS_3238_checkbox = checked Then objTextStream.WriteLine "DHS_3238_checkbox"
			If DHS_2625_checkbox = checked Then objTextStream.WriteLine "DHS_2625_checkbox"
			objTextStream.WriteLine "FORM -a03 - " & case_card_info
			objTextStream.WriteLine "FORM -b03 - " & clt_knows_how_to_use_ebt_card
			objTextStream.WriteLine "FORM -a16 - " & snap_reporting_type
			objTextStream.WriteLine "FORM -b16 - " & next_revw_month
			objTextStream.WriteLine "FORM - 17 - " & confirm_recap_read
			objTextStream.WriteLine "FINAL SUMM - " & case_summary
            objTextStream.WriteLine "SUMM PHONE - " & phone_number_selection
            objTextStream.WriteLine "SUMM MESSG - " & leave_a_message
            objTextStream.WriteLine "SUMM QUEST - " & resident_questions

			For known_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
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
                HH_MEMB_ARRAY(first_name_const, known_membs)&"~"&HH_MEMB_ARRAY(mid_initial, known_membs)&"~"&HH_MEMB_ARRAY(other_names, known_membs)&"~"&HH_MEMB_ARRAY(age, known_membs)&"~"&HH_MEMB_ARRAY(date_of_birth, known_membs)&"~"&HH_MEMB_ARRAY(ssn, known_membs)&"~"&_
                HH_MEMB_ARRAY(ssn_verif, known_membs)&"~"&HH_MEMB_ARRAY(birthdate_verif, known_membs)&"~"&HH_MEMB_ARRAY(gender, known_membs)&"~"&HH_MEMB_ARRAY(race, known_membs)&"~"&HH_MEMB_ARRAY(spoken_lang, known_membs)&"~"&HH_MEMB_ARRAY(written_lang, known_membs)&"~"&_
                HH_MEMB_ARRAY(interpreter, known_membs)&"~"&HH_MEMB_ARRAY(alias_yn, known_membs)&"~"&HH_MEMB_ARRAY(ethnicity_yn, known_membs)&"~"&HH_MEMB_ARRAY(id_verif, known_membs)&"~"&HH_MEMB_ARRAY(rel_to_applcnt, known_membs)&"~"&HH_MEMB_ARRAY(cash_minor, known_membs)&"~"&_
                HH_MEMB_ARRAY(snap_minor, known_membs)&"~"&HH_MEMB_ARRAY(marital_status, known_membs)&"~"&HH_MEMB_ARRAY(spouse_ref, known_membs)&"~"&HH_MEMB_ARRAY(spouse_name, known_membs)&"~"&HH_MEMB_ARRAY(last_grade_completed, known_membs)&"~"&_
                HH_MEMB_ARRAY(citizen, known_membs)&"~"&HH_MEMB_ARRAY(other_st_FS_end_date, known_membs)&"~"&HH_MEMB_ARRAY(in_mn_12_mo, known_membs)&"~"&HH_MEMB_ARRAY(residence_verif, known_membs)&"~"&HH_MEMB_ARRAY(mn_entry_date, known_membs)&"~"&_
                HH_MEMB_ARRAY(former_state, known_membs)&"~"&HH_MEMB_ARRAY(fs_pwe, known_membs)&"~"&HH_MEMB_ARRAY(button_one, known_membs)&"~"&HH_MEMB_ARRAY(button_two, known_membs)&"~"&HH_MEMB_ARRAY(imig_status, known_membs)&"~"&HH_MEMB_ARRAY(clt_has_sponsor, known_membs)&"~"&_
                HH_MEMB_ARRAY(client_verification, known_membs)&"~"&HH_MEMB_ARRAY(client_verification_details, known_membs)&"~"&HH_MEMB_ARRAY(client_notes, known_membs)&"~"&HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)&"~"&race_a_info&"~"&_
                race_b_info&"~"&race_n_info&"~"&race_p_info&"~"&race_w_info&"~"&prog_s_info&"~"&prog_c_info&"~"&prog_e_info&"~"&prog_n_info&"~"&HH_MEMB_ARRAY(ssn_no_space, known_membs)&"~"&HH_MEMB_ARRAY(edrs_msg, known_membs)&"~"&_
                HH_MEMB_ARRAY(edrs_match, known_membs)&"~"&HH_MEMB_ARRAY(edrs_notes, known_membs)&"~"&HH_MEMB_ARRAY(ignore_person, known_membs)&"~"&HH_MEMB_ARRAY(pers_in_maxis, known_membs)&"~"&HH_MEMB_ARRAY(memb_is_caregiver, known_membs)&"~"&_
                HH_MEMB_ARRAY(cash_request_const, known_membs)&"~"&HH_MEMB_ARRAY(hours_per_week_const, known_membs)&"~"&HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(comply_with_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_needed_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(orientation_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_exempt_const, known_membs)&"~"&HH_MEMB_ARRAY(exemption_reason_const, known_membs)&"~"&HH_MEMB_ARRAY(emps_exemption_code_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(choice_form_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_notes, known_membs)&"~"&HH_MEMB_ARRAY(remo_info_const, known_membs)&"~"&HH_MEMB_ARRAY(requires_update, known_membs)&"~"&HH_MEMB_ARRAY(last_const, known_membs)
			Next

			'Close the object so it can be opened again shortly
			objTextStream.Close

			script_run_lowdown = ""
			script_run_lowdown = script_run_lowdown & vbCr & "TIME SPENT - "	& timer - start_time & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "MFIP - ORNT - " & MFIP_orientation_assessed_and_completed & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "MFIP - DWP - " & family_cash_program
            script_run_lowdown = script_run_lowdown & vbCr & "FMCA - 01 - " & famliy_cash_notes & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "RUN BY INTERVIEW TEAM" & run_by_interview_team & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "PROG - CASH - " & cash_other_req_detail
			script_run_lowdown = script_run_lowdown & vbCr & "PROG - SNAP - " & snap_other_req_detail
			script_run_lowdown = script_run_lowdown & vbCr & "PROG - EMER - " & emer_other_req_detail & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "CVR - CMNT - " & additional_application_comments
			script_run_lowdown = script_run_lowdown & vbCr & "CVR - INCM - " & additional_income_comments
			script_run_lowdown = script_run_lowdown & vbCr & "CVR - NOTE - " & cover_letter_interview_notes & vbCr & vbCr

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
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - UPD - " & need_to_update_addr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - RESI - LIV - " & living_situation & vbCr & vbCr

			script_run_lowdown = script_run_lowdown & vbCr & "ADR - HOUS - LIC - " & licensed_facility
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - HOUS - MEA - " & meal_provided
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - HOUS - NAM - " & residence_name_phone & vbCr & vbCr

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
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - CNTY - " & resi_addr_county
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - TEXT - " & send_text
			script_run_lowdown = script_run_lowdown & vbCr & "ADR - EMAL - " & send_email & vbCr & vbCr

            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - ALLYN - " & all_members_listed_yn
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - ALLNT - " & all_members_listed_notes
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - IMNYN - " & all_members_in_MN_yn
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - IMNNT - " & all_members_in_MN_notes
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - PRGYN - " & anyone_pregnant_yn
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - PRGNT - " & anyone_pregnant_notes
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - MILYN - " & anyone_served_yn
            script_run_lowdown = script_run_lowdown & vbCr & "MEMB - MILNT - " & anyone_served_notes & vbCr & vbCr

			For quest = 0 to UBound(FORM_QUESTION_ARRAY)
				Call FORM_QUESTION_ARRAY(quest).add_to_SRL()
			Next

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
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 04 - " & second_signature_detail
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 05 - " & second_signature_person
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 07 - " & client_signed_verbally_yn
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 08 - " & interview_date
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 09 - " & verbal_sig_date
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 10 - " & verbal_sig_time
			script_run_lowdown = script_run_lowdown & vbCr & "SIG - 11 - " & verbal_sig_phone_number & vbCr & vbCr

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

			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 01 - " & exp_det_income
			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 02 - " & exp_det_assets
			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 03 - " & exp_det_housing
			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 04 - " & exp_det_utilities
			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 05 - " & exp_det_notes
			script_run_lowdown = script_run_lowdown & vbCr & "SIMP EXPDET - 06 - " & expedited_viewed
            If verif_snap_checkbox = checked 	then script_run_lowdown = script_run_lowdown & vbCr & "verif_snap_checkbox - CHECKED"
			If heat_exp_checkbox = checked 		then script_run_lowdown = script_run_lowdown & vbCr & "heat_exp_checkbox - CHECKED"
			If ac_exp_checkbox = checked 		then script_run_lowdown = script_run_lowdown & vbCr & "ac_exp_checkbox - CHECKED"
			If electric_exp_checkbox = checked 	then script_run_lowdown = script_run_lowdown & vbCr & "electric_exp_checkbox - CHECKED"
			If phone_exp_checkbox = checked 	then script_run_lowdown = script_run_lowdown & vbCr & "phone_exp_checkbox - CHECKED"
			If none_exp_checkbox = checked 		then script_run_lowdown = script_run_lowdown & vbCr & "none_exp_checkbox - CHECKED" & vbCr & vbCr

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

			'Emergency questions
			script_run_lowdown = script_run_lowdown & vbCr & "EMER EXP - " & resident_emergency_yn
			script_run_lowdown = script_run_lowdown & vbCr & "EMER TYPE - " & emergency_type
			script_run_lowdown = script_run_lowdown & vbCr & "EMER DISCUSSION - " & emergency_discussion
			script_run_lowdown = script_run_lowdown & vbCr & "EMER AMOUNT - " & emergency_amount
			script_run_lowdown = script_run_lowdown & vbCr & "EMER DEADLINE - " & emergency_deadline & vbCr & vbCr

			'Interview HSR E&T Questions
			script_run_lowdown = script_run_lowdown & vbCr & "SUMM INTNOW - " & interested_in_job_assistance_now
			script_run_lowdown = script_run_lowdown & vbCr & "SUMM NAMENOW - " & interested_names_now
			script_run_lowdown = script_run_lowdown & vbCr & "SUMM INTFUT - " & interested_in_job_assistance_future
			script_run_lowdown = script_run_lowdown & vbCr & "SUMM NAMEFUT - " & interested_names_future & vbCr & vbCr

			'R&R
            If DHS_4163_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_4163_checkbox - CHECKED"
            If DHS_3315A_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3315A_checkbox - CHECKED"
            If DHS_3979_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3979_checkbox - CHECKED"
            If DHS_2759_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2759_checkbox - CHECKED"
            If DHS_3353_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3353_checkbox - CHECKED"
            If DHS_2920_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2920_checkbox - CHECKED"
            If DHS_3477_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3477_checkbox - CHECKED"
            If DHS_4133_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_4133_checkbox - CHECKED"
            If DHS_2647_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2647_checkbox - CHECKED"
            If DHS_2929_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2929_checkbox - CHECKED"
            If DHS_3323_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3323_checkbox - CHECKED"
            If DHS_3393_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3393_checkbox - CHECKED"
            If DHS_3163B_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3163B_checkbox - CHECKED"
            If DHS_2338_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2338_checkbox - CHECKED"
            If DHS_5561_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_5561_checkbox - CHECKED"
            If DHS_2961_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2961_checkbox - CHECKED"
            If DHS_2887_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2887_checkbox - CHECKED"
            If DHS_3238_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_3238_checkbox - CHECKED"
            If DHS_2625_checkbox = checked then script_run_lowdown = script_run_lowdown & vbCr & "DHS_2625_checkbox - CHECKED" & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -a03 - " & case_card_info
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -b03 - " & clt_knows_how_to_use_ebt_card
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -a16 - " & snap_reporting_type
			script_run_lowdown = script_run_lowdown & vbCr & "FORM -b16 - " & next_revw_month
			script_run_lowdown = script_run_lowdown & vbCr & "FORM - 17 - " & confirm_recap_read & vbCr & vbCr
			script_run_lowdown = script_run_lowdown & vbCr & "FINAL SUMM - " & case_summary & vbCr & vbCr
            script_run_lowdown = script_run_lowdown & vbCr & "SUMM PHONE - " & phone_number_selection
            script_run_lowdown = script_run_lowdown & vbCr & "SUMM MESSG - " & leave_a_message
            script_run_lowdown = script_run_lowdown & vbCr & "SUMM QUEST - " & resident_questions & vbCr & vbCr


			For known_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
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
                HH_MEMB_ARRAY(first_name_const, known_membs)&"~"&HH_MEMB_ARRAY(mid_initial, known_membs)&"~"&HH_MEMB_ARRAY(other_names, known_membs)&"~"&HH_MEMB_ARRAY(age, known_membs)&"~"&HH_MEMB_ARRAY(date_of_birth, known_membs)&"~"&HH_MEMB_ARRAY(ssn, known_membs)&"~"&_
                HH_MEMB_ARRAY(ssn_verif, known_membs)&"~"&HH_MEMB_ARRAY(birthdate_verif, known_membs)&"~"&HH_MEMB_ARRAY(gender, known_membs)&"~"&HH_MEMB_ARRAY(race, known_membs)&"~"&HH_MEMB_ARRAY(spoken_lang, known_membs)&"~"&HH_MEMB_ARRAY(written_lang, known_membs)&"~"&_
                HH_MEMB_ARRAY(interpreter, known_membs)&"~"&HH_MEMB_ARRAY(alias_yn, known_membs)&"~"&HH_MEMB_ARRAY(ethnicity_yn, known_membs)&"~"&HH_MEMB_ARRAY(id_verif, known_membs)&"~"&HH_MEMB_ARRAY(rel_to_applcnt, known_membs)&"~"&HH_MEMB_ARRAY(cash_minor, known_membs)&"~"&_
                HH_MEMB_ARRAY(snap_minor, known_membs)&"~"&HH_MEMB_ARRAY(marital_status, known_membs)&"~"&HH_MEMB_ARRAY(spouse_ref, known_membs)&"~"&HH_MEMB_ARRAY(spouse_name, known_membs)&"~"&HH_MEMB_ARRAY(last_grade_completed, known_membs)&"~"&_
                HH_MEMB_ARRAY(citizen, known_membs)&"~"&HH_MEMB_ARRAY(other_st_FS_end_date, known_membs)&"~"&HH_MEMB_ARRAY(in_mn_12_mo, known_membs)&"~"&HH_MEMB_ARRAY(residence_verif, known_membs)&"~"&HH_MEMB_ARRAY(mn_entry_date, known_membs)&"~"&_
                HH_MEMB_ARRAY(former_state, known_membs)&"~"&HH_MEMB_ARRAY(fs_pwe, known_membs)&"~"&HH_MEMB_ARRAY(button_one, known_membs)&"~"&HH_MEMB_ARRAY(button_two, known_membs)&"~"&HH_MEMB_ARRAY(imig_status, known_membs)&"~"&HH_MEMB_ARRAY(clt_has_sponsor, known_membs)&"~"&_
                HH_MEMB_ARRAY(client_verification, known_membs)&"~"&HH_MEMB_ARRAY(client_verification_details, known_membs)&"~"&HH_MEMB_ARRAY(client_notes, known_membs)&"~"&HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)&"~"&race_a_info&"~"&_
                race_b_info&"~"&race_n_info&"~"&race_p_info&"~"&race_w_info&"~"&prog_s_info&"~"&prog_c_info&"~"&prog_e_info&"~"&prog_n_info&"~"&HH_MEMB_ARRAY(ssn_no_space, known_membs)&"~"&HH_MEMB_ARRAY(edrs_msg, known_membs)&"~"&_
                HH_MEMB_ARRAY(edrs_match, known_membs)&"~"&HH_MEMB_ARRAY(edrs_notes, known_membs)&"~"&HH_MEMB_ARRAY(ignore_person, known_membs)&"~"&HH_MEMB_ARRAY(pers_in_maxis, known_membs)&"~"&HH_MEMB_ARRAY(memb_is_caregiver, known_membs)&"~"&_
                HH_MEMB_ARRAY(cash_request_const, known_membs)&"~"&HH_MEMB_ARRAY(hours_per_week_const, known_membs)&"~"&HH_MEMB_ARRAY(exempt_from_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(comply_with_ed_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_needed_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(orientation_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_exempt_const, known_membs)&"~"&HH_MEMB_ARRAY(exemption_reason_const, known_membs)&"~"&HH_MEMB_ARRAY(emps_exemption_code_const, known_membs)&"~"&_
                HH_MEMB_ARRAY(choice_form_done_const, known_membs)&"~"&HH_MEMB_ARRAY(orientation_notes, known_membs)&"~"&HH_MEMB_ARRAY(remo_info_const, known_membs)&"~"&HH_MEMB_ARRAY(requires_update, known_membs)&"~"&HH_MEMB_ARRAY(last_const, known_membs) & vbCr & vbCr
			Next

			'Since the file was new, we can simply exit the function
			exit function
		End if
	End with
end function

function restore_your_work(vars_filled, membs_found)
'this function looks to see if a txt file exists for the case that is being run to pull already known variables back into the script from a previous run

	'Now determines name of file
	save_your_work_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"

	With (CreateObject("Scripting.FileSystemObject"))

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(save_your_work_path) = True then

			pull_variables = MsgBox("It appears there is information saved for this case from a previous run of this script." & vbCr & vbCr & "THE FORM SELECTION CANNOT BE CHANGED." & vbCr & "The script will load the form selected in the original run. If the form needs to be changed select NO to restoring details from the previous run." & vbCr & vbCr & "Would you like to restore the details from this previous run?", vbQuestion + vbYesNo, "Restore Detail from Previous Run")

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

					If left(text_line, 11) = "VERB - CASH" Then cash_verbal_request = Mid(text_line, 15)
					If left(text_line, 11) = "VERB - GRHS" Then grh_verbal_request = Mid(text_line, 15)
					If left(text_line, 11) = "VERB - SNAP" Then snap_verbal_request = Mid(text_line, 15)
					If left(text_line, 11) = "VERB - EMER" Then emer_verbal_request = Mid(text_line, 15)
					If left(text_line, 11) = "WTDR - CASH" Then cash_verbal_withdraw = Mid(text_line, 15)
					If left(text_line, 11) = "WTDR - GRHS" Then grh_verbal_withdraw = Mid(text_line, 15)
					If left(text_line, 11) = "WTDR - SNAP" Then snap_verbal_withdraw = Mid(text_line, 15)
					If left(text_line, 11) = "WTDR - EMER" Then emer_verbal_withdraw = Mid(text_line, 15)
					If left(text_line, 17) = "CASH PROG CHECKED" Then CASH_on_CAF_checkbox = checked
					If left(text_line, 17) = "GRHS PROG CHECKED" Then GRH_on_CAF_checkbox = checked
					If left(text_line, 17) = "SNAP PROG CHECKED" Then SNAP_on_CAF_checkbox = checked
					If left(text_line, 17) = "EMER PROG CHECKED" Then EMER_on_CAF_checkbox = checked

					If left(text_line, 11) = "CASH - TYPE" Then type_of_cash = Mid(text_line, 15)
					If left(text_line, 11) = "PROC - CASH" Then the_process_for_cash = Mid(text_line, 15)
					If left(text_line, 11) = "CASH - RVMO" Then next_cash_revw_mo = Mid(text_line, 15)
					If left(text_line, 11) = "CASH - RVYR" Then next_cash_revw_yr = Mid(text_line, 15)

					If left(text_line, 11) = "PROC - SNAP" Then the_process_for_snap = Mid(text_line, 15)
					If left(text_line, 11) = "SNAP - RVMO" Then next_snap_revw_mo = Mid(text_line, 15)
					If left(text_line, 11) = "SNAP - RVYR" Then next_snap_revw_yr = Mid(text_line, 15)

					If left(text_line, 11) = "PROC - GRHS" Then the_process_for_grh = Mid(text_line, 15)
					If left(text_line, 11) = "GRHS - RVMO" Then next_grh_revw_mo = Mid(text_line, 15)
					If left(text_line, 11) = "GRHS - RVYR" Then next_grh_revw_yr = Mid(text_line, 15)

					If left(text_line, 11) = "EMER - TYPE" Then type_of_emer = Mid(text_line, 15)
					If left(text_line, 11) = "PROC - EMER" Then the_process_for_emer = Mid(text_line, 15)

					If left(text_line, 11) = "PROG - NOTE" Then program_request_notes = Mid(text_line, 15)
					If left(text_line, 11) = "VERB - NOTE" Then verbal_request_notes = Mid(text_line, 15)

			        If left(text_line, 10) = "CVR - CMNT" Then additional_application_comments = Mid(text_line, 14)
			        If left(text_line, 10) = "CVR - INCM" Then additional_income_comments = Mid(text_line, 14)
			        If left(text_line, 10) = "CVR - NOTE" Then cover_letter_interview_notes = Mid(text_line, 14)


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
						If mid(text_line, 7, 10) = "RESI - UPD" Then need_to_update_addr = MID(text_line, 20)

						If mid(text_line, 7, 10) = "MAIL - STR" Then mail_addr_street_full = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - CIT" Then mail_addr_city = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - STA" Then mail_addr_state = MID(text_line, 20)
						If mid(text_line, 7, 10) = "MAIL - ZIP" Then mail_addr_zip = MID(text_line, 20)


						If mid(text_line, 7, 10) = "HOUS - LIC" Then licensed_facility = MID(text_line, 20)
						If mid(text_line, 7, 10) = "HOUS - MEA" Then meal_provided = MID(text_line, 20)
						If mid(text_line, 7, 10) = "HOUS - NAM" Then residence_name_phone = MID(text_line, 20)

						If mid(text_line, 7, 10) = "PHON - NON" Then phone_one_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TON" Then phone_one_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTW" Then phone_two_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTW" Then phone_two_type = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - NTH" Then phone_three_number = MID(text_line, 20)
						If mid(text_line, 7, 10) = "PHON - TTH" Then phone_three_type = MID(text_line, 20)

						If mid(text_line, 7, 4) = "DATE" Then address_change_date = MID(text_line, 14)
						If mid(text_line, 7, 4) = "CNTY" Then resi_addr_county = MID(text_line, 14)
						If mid(text_line, 7, 4) = "TEXT" Then send_text = MID(text_line, 14)
						If mid(text_line, 7, 4) = "EMAL" Then send_email = MID(text_line, 14)
					End If

                    If left(text_line, 12) = "MEMB - ALLYN" Then all_members_listed_yn = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - ALLNT" Then all_members_listed_notes = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - IMNYN" Then all_members_in_MN_yn = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - IMNNT" Then all_members_in_MN_notes = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - PRGYN" Then anyone_pregnant_yn = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - PRGNT" Then anyone_pregnant_notes = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - MILYN" Then anyone_served_yn = Mid(text_line, 16)
                    If left(text_line, 12) = "MEMB - MILNT" Then anyone_served_notes = Mid(text_line, 16)

					If left(text_line, 3) = "PWE" Then pwe_selection = Mid(text_line, 7)

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
					If left(text_line, 8) = "SIG - 04" Then second_signature_detail = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 05" Then second_signature_person = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 07" Then client_signed_verbally_yn = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 08" Then interview_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 09" Then verbal_sig_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 10" Then verbal_sig_time = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 11" Then verbal_sig_phone_number	 = Mid(text_line, 12)

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

					If left(text_line, 16) = "SIMP EXPDET - 01" Then exp_det_income = Mid(text_line, 20)
					If left(text_line, 16) = "SIMP EXPDET - 02" Then exp_det_assets = Mid(text_line, 20)
					If left(text_line, 16) = "SIMP EXPDET - 03" Then exp_det_housing = Mid(text_line, 20)
					If left(text_line, 16) = "SIMP EXPDET - 04" Then exp_det_utilities = Mid(text_line, 20)
					If left(text_line, 16) = "SIMP EXPDET - 05" Then exp_det_notes = Mid(text_line, 20)
					If left(text_line, 16) = "SIMP EXPDET - 06" Then read_exp_view = Mid(text_line, 20)
					If Ucase(read_exp_view) = "TRUE" Then expedited_viewed = True
					If UCase(read_exp_view) = "FALSE" Then expedited_viewed = False

					If text_line = "verif_snap_checkbox" Then verif_snap_checkbox = checked
					If text_line = "heat_exp_checkbox" Then heat_exp_checkbox = checked
					If text_line = "ac_exp_checkbox" Then ac_exp_checkbox = checked
					If text_line = "electric_exp_checkbox" Then electric_exp_checkbox = checked
					If text_line = "phone_exp_checkbox" Then phone_exp_checkbox = checked
					If text_line = "none_exp_checkbox" Then none_exp_checkbox = checked

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

					'Emergency questions
					If left(text_line, 11) = "EMER EXP - " Then resident_emergency_yn = Mid(text_line, 12)
					If left(text_line, 12) = "EMER TYPE - " Then emergency_type = Mid(text_line, 13)
					If left(text_line, 18) = "EMER DISCUSSION - " Then emergency_discussion = Mid(text_line, 19)
					If left(text_line, 14) = "EMER AMOUNT - " Then emergency_amount = Mid(text_line, 15)
					If left(text_line, 16) = "EMER DEADLINE - " Then emergency_deadline = Mid(text_line, 17)


					'Interview HSR E&T Questions
					If left(text_line, 14) = "SUMM INTNOW - " Then interested_in_job_assistance_now = Mid(text_line, 15)
					If left(text_line, 15) = "SUMM NAMENOW - " Then interested_names_now = Mid(text_line, 16)
					If left(text_line, 14) = "SUMM INTFUT - " Then interested_in_job_assistance_future = Mid(text_line, 15)
					If left(text_line, 15) = "SUMM NAMEFUT - " Then interested_names_future = Mid(text_line, 16)

					'R&R
					If text_line = "DHS_4163_checkbox" Then DHS_4163_checkbox = checked
					If text_line = "DHS_3315A_checkbox" Then DHS_3315A_checkbox = checked
					If text_line = "DHS_3979_checkbox" Then DHS_3979_checkbox = checked
					If text_line = "DHS_2759_checkbox" Then DHS_2759_checkbox = checked
					If text_line = "DHS_3353_checkbox" Then DHS_3353_checkbox = checked
					If text_line = "DHS_2920_checkbox" Then DHS_2920_checkbox = checked
					If text_line = "DHS_3477_checkbox" Then DHS_3477_checkbox = checked
					If text_line = "DHS_4133_checkbox" Then DHS_4133_checkbox = checked
					If text_line = "DHS_2647_checkbox" Then DHS_2647_checkbox = checked
					If text_line = "DHS_2929_checkbox" Then DHS_2929_checkbox = checked
					If text_line = "DHS_3323_checkbox" Then DHS_3323_checkbox = checked
					If text_line = "DHS_3393_checkbox" Then DHS_3393_checkbox = checked
					If text_line = "DHS_3163B_checkbox" Then DHS_3163B_checkbox = checked
					If text_line = "DHS_2338_checkbox" Then DHS_2338_checkbox = checked
					If text_line = "DHS_5561_checkbox" Then DHS_5561_checkbox = checked
					If text_line = "DHS_2961_checkbox" Then DHS_2961_checkbox = checked
					If text_line = "DHS_2887_checkbox" Then DHS_2887_checkbox = checked
					If text_line = "DHS_3238_checkbox" Then DHS_3238_checkbox = checked
					If text_line = "DHS_2625_checkbox" Then DHS_2625_checkbox = checked
					If left(text_line, 9) = "FORM -a03" Then case_card_info = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -b03" Then clt_knows_how_to_use_ebt_card = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -a16" Then snap_reporting_type = Mid(text_line, 13)
					If left(text_line, 9) = "FORM -b16" Then next_revw_month = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 17" Then confirm_recap_read = Mid(text_line, 13)
					If left(text_line, 10) = "FINAL SUMM" Then case_summary = Mid(text_line, 14)
                    If left(text_line, 10) = "SUMM PHONE" Then phone_number_selection = Mid(text_line, 14)
                    If left(text_line, 10) = "SUMM MESSG" Then leave_a_message = Mid(text_line, 14)
                    If left(text_line, 10) = "SUMM QUEST" Then resident_questions = Mid(text_line, 14)

					If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)


					If left(text_line, 3) = "ARR" Then
						If MID(text_line, 7, 13) = "HH_MEMB_ARRAY" Then
							array_info = Mid(text_line, 23)
							array_info = split(array_info, "~")
                            If array_info(0) <> "" Then
                                membs_found = True
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
                                ' 36
                                HH_MEMB_ARRAY(imig_status, known_membs)					= array_info(36)

                                HH_MEMB_ARRAY(clt_has_sponsor, known_membs)				= array_info(37)
                                HH_MEMB_ARRAY(client_verification, known_membs)			= array_info(38)
                                HH_MEMB_ARRAY(client_verification_details, known_membs)	= array_info(39)
                                HH_MEMB_ARRAY(client_notes, known_membs)				= array_info(40)
                                HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)		= array_info(41)
                                If array_info(42) = "YES" Then HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked
                                If array_info(43) = "YES" Then HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked
                                If array_info(44) = "YES" Then HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked
                                If array_info(45) = "YES" Then HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked
                                If array_info(46) = "YES" Then HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked
                                If array_info(47) = "YES" Then HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked
                                If array_info(48) = "YES" Then HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked
                                If array_info(49) = "YES" Then HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked
                                If array_info(50) = "YES" Then HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked
                                HH_MEMB_ARRAY(ssn_no_space, known_membs)				= array_info(51)
                                HH_MEMB_ARRAY(edrs_msg, known_membs)					= array_info(52)
                                HH_MEMB_ARRAY(edrs_match, known_membs)					= array_info(53)
                                HH_MEMB_ARRAY(edrs_notes, known_membs) 					= array_info(54)

                                HH_MEMB_ARRAY(ignore_person, known_membs) 			= array_info(55)
                                HH_MEMB_ARRAY(pers_in_maxis, known_membs) 			= array_info(56)

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
                                If UBound(array_info) = 69 Then
                                    HH_MEMB_ARRAY(last_const, known_membs)              = array_info(69)
                                Else
                                    HH_MEMB_ARRAY(remo_info_const, known_membs)         = array_info(69)
                                    HH_MEMB_ARRAY(requires_update, known_membs)         = array_info(70)
                                    HH_MEMB_ARRAY(last_const, known_membs)				= array_info(71)
                                End If

                                known_membs = known_membs + 1
                            End If
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
	If anyone_pregnant_yn = "" Then anyone_pregnant_yn = exp_pregnant_yn
    If signature_detail = "Accepted Verbally" or second_signature_detail = "Accepted Verbally" Then
		If verbal_sig_date = "" Then verbal_sig_date = date & ""
		If verbal_sig_time = "" Then
			time_hr = DatePart("h", time)
			time_min = DatePart("n", time)
			verbal_sig_time = time_hr & ":" & time_min
			verbal_sig_time = FormatDateTime(verbal_sig_time, 3)
			verbal_sig_time = replace(verbal_sig_time, ":00 ", " ")
		End If
	End If

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), "")
		If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
			If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), "")
			If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_HOUSING_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_UTILITIES_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_ASSETS_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_MSA_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "stwk" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_INFO_ARRAY(form_second_ans_const, quest), TEMP_STWK_ARRAY)
		End If
	Next

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

        HH_MEMB_ARRAY(last_name_const, the_memb) = UCase(trim(HH_MEMB_ARRAY(last_name_const, the_memb)))
        HH_MEMB_ARRAY(first_name_const, the_memb) = UCase(trim(HH_MEMB_ARRAY(first_name_const, the_memb)))
		HH_MEMB_ARRAY(mid_initial, the_memb) = UCase(trim(HH_MEMB_ARRAY(mid_initial, the_memb)))

		HH_MEMB_ARRAY(full_name_const, the_memb) = HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb)
        If NOT IsNumeric(HH_MEMB_ARRAY(age, the_memb)) and IsDate(HH_MEMB_ARRAY(date_of_birth, the_memb)) Then HH_MEMB_ARRAY(age, the_memb) = DateDiff("yyyy", CDate(HH_MEMB_ARRAY(date_of_birth, the_memb)), Date)

        race_string = ""
        If HH_MEMB_ARRAY(race_a_checkbox, the_memb) = checked Then race_string = race_string & "Asian~"
        If HH_MEMB_ARRAY(race_b_checkbox, the_memb) = checked Then race_string = race_string & "Black Or African Amer~"
        If HH_MEMB_ARRAY(race_n_checkbox, the_memb) = checked Then race_string = race_string & "Amer Indn Or Alaskan Native~"
        If HH_MEMB_ARRAY(race_p_checkbox, the_memb) = checked Then race_string = race_string & "Pacific Is Or Native Hawaii~"
        If HH_MEMB_ARRAY(race_w_checkbox, the_memb) = checked Then race_string = race_string & "White~"
        If right(race_string, 1) = "~" Then race_string = left(race_string, len(race_string) - 1)
        If InStr(race_string, "~") > 0 Then race_string = "Multiple Races"
        If HH_MEMB_ARRAY(race, the_memb) = "Multiple Races" and race_string = "" Then race_string = "Multiple Races"
        If race_string = "" Then race_string = "Unable To Determine"
        HH_MEMB_ARRAY(race, the_memb) = race_string

        If the_memb = 0 Then
            HH_arrived_date = HH_MEMB_ARRAY(mn_entry_date, the_memb)
            HH_arrived_place = HH_MEMB_ARRAY(former_state, the_memb)
        Else
            If HH_arrived_date <> HH_MEMB_ARRAY(mn_entry_date, the_memb) or HH_arrived_place <> HH_MEMB_ARRAY(former_state, the_memb)  Then
                HH_arrived_date = ""
                HH_arrived_place = ""
            End If
        End If

        HH_MEMB_ARRAY(requires_update, the_memb) = False
        If HH_MEMB_ARRAY(rel_to_applcnt, the_memb) = "01 Self" and (HH_MEMB_ARRAY(id_verif, the_memb) = "__" or HH_MEMB_ARRAY(id_verif, the_memb) = "NO - No Ver Prvd") Then
            HH_MEMB_ARRAY(requires_update, the_memb) = True
        End If

        ssn_info_valid = True
        If trim(HH_MEMB_ARRAY(ssn, the_memb)) = "" Then
            ssn_info_valid = False
            If HH_MEMB_ARRAY(ssn_verif, the_memb) = "A - SSN Applied For" Then ssn_info_valid = True
            If HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - Member Does Not Have SSN" Then ssn_info_valid = True
        End If
        If HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - SSN Not Provided" Then ssn_info_valid = False
        If HH_MEMB_ARRAY(none_req_checkbox, the_memb) = checked Then ssn_info_valid = True
        If ssn_info_valid = False Then HH_MEMB_ARRAY(requires_update, the_memb) = True

        If HH_MEMB_ARRAY(citizen, the_memb) = "No" and HH_MEMB_ARRAY(none_req_checkbox, the_memb) = unchecked Then
            If trim(HH_MEMB_ARRAY(imig_status, the_memb)) = "" Then
                HH_MEMB_ARRAY(requires_update, the_memb) = True
            End If
            If (HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "?" or HH_MEMB_ARRAY(clt_has_sponsor, the_memb) = "") and HH_MEMB_ARRAY(none_req_checkbox, the_memb) = unchecked Then
                HH_MEMB_ARRAY(requires_update, the_memb) = True
            End If
        End If
	next
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

	'PHONE NUMBER BUT NO PHONE EXPENSE
	disc_yes_phone_no_expense_confirmation = trim(disc_yes_phone_no_expense_confirmation)
	disc_no_phone_yes_expense_confirmation = trim(disc_no_phone_yes_expense_confirmation)
	phone_details = trim(phone_details)
	disc_yes_phone_no_expense = "N/A"
	disc_no_phone_yes_expense = "N/A"

	phone_number_entered = True
	If phone_one_number = "" AND phone_two_number = "" AND phone_three_number = "" Then phone_number_entered = False

    If expedited_screening_on_form Then
        If caf_exp_pay_phone_checkbox = unchecked AND phone_number_entered = True Then disc_yes_phone_no_expense = "EXISTS"
        If caf_exp_pay_phone_checkbox = checked AND phone_number_entered = False Then disc_no_phone_yes_expense = "EXISTS"
    End If

	rent_indicated = False
	rent_summary = ""
	utility_summary = ""
	disc_utility_amounts = "N/A"

	For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
		If FORM_QUESTION_ARRAY(each_question).detail_source = "shel-hest" Then
            If expedited_screening_on_form Then
                If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = checked Then utility_summary = utility_summary & ", Heat/AC"
                If FORM_QUESTION_ARRAY(each_question).electric_checkbox = checked Then utility_summary = utility_summary & ", Electric"
                If FORM_QUESTION_ARRAY(each_question).phone_checkbox = checked Then utility_summary = utility_summary & ", Phone"
                If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = unchecked AND caf_exp_pay_heat_checkbox = checked 		Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = unchecked AND caf_exp_pay_ac_checkbox = checked 			Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).electric_checkbox = unchecked AND caf_exp_pay_electricity_checkbox = checked 	Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).phone_checkbox = unchecked 	AND caf_exp_pay_phone_checkbox = checked 		Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = checked   AND caf_exp_pay_heat_checkbox = unchecked 		Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = checked   AND caf_exp_pay_ac_checkbox = unchecked 			Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).electric_checkbox = checked   AND caf_exp_pay_electricity_checkbox = unchecked 	Then disc_utility_amounts = "EXISTS"
                If FORM_QUESTION_ARRAY(each_question).phone_checkbox = checked 	    AND caf_exp_pay_phone_checkbox = unchecked 		Then disc_utility_amounts = "EXISTS"
                If caf_exp_pay_none_checkbox = checked Then
                    If FORM_QUESTION_ARRAY(each_question).heat_air_checkbox = checked 	Then disc_utility_amounts = "EXISTS"
                    If FORM_QUESTION_ARRAY(each_question).electric_checkbox = checked 	Then disc_utility_amounts = "EXISTS"
                    If FORM_QUESTION_ARRAY(each_question).phone_checkbox = checked 		Then disc_utility_amounts = "EXISTS"
                End If
            End If
			If FORM_QUESTION_ARRAY(each_question).phone_checkbox = checked AND phone_number_entered = False Then disc_no_phone_yes_expense = "EXISTS"
			If FORM_QUESTION_ARRAY(each_question).phone_checkbox = unchecked AND phone_number_entered = True Then disc_yes_phone_no_expense = "EXISTS"
			If trim(FORM_QUESTION_ARRAY(each_question).housing_payment) <> "" Then
				rent_summary = rent_summary & "/Housing Payment: " &FORM_QUESTION_ARRAY(each_question).housing_payment
				rent_indicated = True
			End If
		End If

		If FORM_QUESTION_ARRAY(each_question).info_type = "housing" Then
			If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
				For each_shel = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
					If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_shel) = "Yes" Then
						rent_indicated = True
						rent_summary = rent_summary & "/" & FORM_QUESTION_ARRAY(each_question).item_info_list(each_shel)
					End If
				Next
			End If
		End If

		If FORM_QUESTION_ARRAY(each_question).info_type = "utilities" Then
			If FORM_QUESTION_ARRAY(each_question).answer_is_array = true Then
				For each_util = 0 to UBound(FORM_QUESTION_ARRAY(each_question).item_info_list)
					If FORM_QUESTION_ARRAY(each_question).item_ans_list(each_util) = "Yes" Then
                        If expedited_screening_on_form Then
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) <> "Water/Sewer" AND FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) <> "Garbage" Then utility_summary = utility_summary & ", " & FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util)
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Heat" 	    AND caf_exp_pay_none_checkbox = checked             Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "AC" 	    AND caf_exp_pay_none_checkbox = checked             Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Electric"   AND caf_exp_pay_none_checkbox = checked             Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone"      AND caf_exp_pay_none_checkbox = checked             Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Heat" 	    AND caf_exp_pay_heat_checkbox = unchecked           Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "AC" 	    AND caf_exp_pay_ac_checkbox = unchecked             Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Electric"   AND caf_exp_pay_electricity_checkbox = unchecked    Then disc_utility_amounts = "EXISTS"
                            If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone"      AND caf_exp_pay_phone_checkbox = unchecked          Then disc_utility_amounts = "EXISTS"
                        End If
                        If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone" AND phone_number_entered = False Then disc_no_phone_yes_expense = "EXISTS"
					Else
                        If expedited_screening_on_form Then
                            If caf_exp_pay_heat_checkbox = checked AND 			FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Heat" 	    Then disc_utility_amounts = "EXISTS"
                            If caf_exp_pay_ac_checkbox = checked AND 			FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "AC" 	    Then disc_utility_amounts = "EXISTS"
                            If caf_exp_pay_electricity_checkbox = checked AND 	FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Electric" 	Then disc_utility_amounts = "EXISTS"
                            If caf_exp_pay_phone_checkbox = checked AND 		FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone" 	Then disc_utility_amounts = "EXISTS"
                        End If
						If FORM_QUESTION_ARRAY(each_question).item_note_info_list(each_util) = "Phone" AND phone_number_entered = True Then disc_yes_phone_no_expense = "EXISTS"
					End If
				Next
			End If
		End If
	Next
    If left(utility_summary, 2) = ", " Then utility_summary = right(utility_summary, len(utility_summary) - 2)

	If disc_yes_phone_no_expense <> "N/A" Then
		If disc_yes_phone_no_expense_confirmation <> "" and disc_yes_phone_no_expense_confirmation <> "Select or Type" Then disc_yes_phone_no_expense = "RESOLVED"
	Else
		disc_yes_phone_no_expense_confirmation = ""
	End If
	If disc_no_phone_yes_expense <> "N/A" Then
		If disc_no_phone_yes_expense_confirmation <> "" and disc_no_phone_yes_expense_confirmation <> "Select or Type" Then disc_no_phone_yes_expense = "RESOLVED"
	Else
		disc_no_phone_yes_expense_confirmation = ""
	End If

	If left(rent_summary, 1) = "/" Then rent_summary = right(rent_summary, len(rent_summary) - 1)
	If rent_summary = "" Then rent_summary = "None Indicated"

	If CAF1_rent_indicated <> rent_indicated Then disc_rent_amounts = "EXISTS"
	If CAF1_rent_indicated = rent_indicated Then disc_rent_amounts = "N/A"

	If disc_rent_amounts <> "N/A" Then
		If disc_rent_amounts_confirmation <> "" and disc_rent_amounts_confirmation <> "Select or Type" Then disc_rent_amounts = "RESOLVED"
	Else
		disc_rent_amounts_confirmation = ""
	End If

	disc_utility_caf_1_summary = ""
	If caf_exp_pay_heat_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Heat"
	If caf_exp_pay_ac_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", AC"
	If caf_exp_pay_electricity_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Electricity"
	If caf_exp_pay_phone_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", Phone"
	If caf_exp_pay_none_checkbox = checked Then disc_utility_caf_1_summary = disc_utility_caf_1_summary & ", NONE"
	If left(disc_utility_caf_1_summary, 1) = "," Then disc_utility_caf_1_summary = right(disc_utility_caf_1_summary, len(disc_utility_caf_1_summary) - 2)

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

				For quest = 0 to UBound(FORM_QUESTION_ARRAY)
					If FORM_QUESTION_ARRAY(quest).verif_status = "Requested" Then
						Text 10, y_pos, 500, 10, FORM_QUESTION_ARRAY(quest).verif_verbiage & " - " & FORM_QUESTION_ARRAY(quest).verif_notes
						y_pos = y_pos + 15
						grp_len = grp_len + 15
					End If
					If FORM_QUESTION_ARRAY(quest).detail_array_exists = True Then
						For each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
							If FORM_QUESTION_ARRAY(quest).detail_verif_status(each_item) = "Requested" Then
								item_information = ""
								If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " at " & FORM_QUESTION_ARRAY(quest).detail_business(each_item)
								If FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " of " & FORM_QUESTION_ARRAY(quest).detail_type(each_item)
								If FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " of " & FORM_QUESTION_ARRAY(quest).detail_type(each_item)
								If FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then item_information = FORM_QUESTION_ARRAY(quest).detail_type(each_item)
								If FORM_QUESTION_ARRAY(quest).detail_source = "expense" Then item_information = "expense paid by " & FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item)
								If FORM_QUESTION_ARRAY(quest).detail_source = "changes" Then item_information = "change"
								If FORM_QUESTION_ARRAY(quest).detail_source = "winnings" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " winnings"

								Text 10, y_pos, 500, 10, FORM_QUESTION_ARRAY(quest).verif_verbiage & " - " & item_information & ". Details: " & FORM_QUESTION_ARRAY(quest).detail_verif_notes(each_item)
								y_pos = y_pos + 15
								grp_len = grp_len + 15
							End If
						Next
					End If
				Next

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
				  Text 10, 20, 490, 10, "Note: After you press 'Update' or 'Return to Dialog' the information from the boxes will be added to the list of verification and the boxes will be 'unchecked'."
				  ButtonGroup ButtonPressed
					PushButton 485, 10, 50, 15, "Update", fill_button
			  End If


              ButtonGroup ButtonPressed
			  	PushButton 545, 365, 60, 15, "Return to Dialog", return_to_dialog_button
				PushButton 5, 365, 125, 15, "Clear ALL Requested Verifications", clear_verifs_btn
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
			If ButtonPressed = clear_verifs_btn Then
				verif_err_msg = "LOOP" & verif_err_msg
				verifs_selected = ""

				For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
					If HH_MEMB_ARRAY(client_verification, the_members) = "Requested" Then
						HH_MEMB_ARRAY(client_verification, the_members) = ""
						HH_MEMB_ARRAY(client_verification_details, the_members) = ""
					End If
				Next

				For quest = 0 to UBound(FORM_QUESTION_ARRAY)
					If FORM_QUESTION_ARRAY(quest).verif_status = "Requested" Then
						FORM_QUESTION_ARRAY(quest).verif_status = ""
						FORM_QUESTION_ARRAY(quest).verif_notes = ""
					End If
					If FORM_QUESTION_ARRAY(quest).detail_array_exists = True Then
						For each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
							If FORM_QUESTION_ARRAY(quest).detail_verif_status(each_item) = "Requested" Then
								FORM_QUESTION_ARRAY(quest).detail_verif_status(each_item) = ""
								FORM_QUESTION_ARRAY(quest).detail_verif_notes(each_item) = ""
							End If
						Next
					End If
				Next
			End If
        Loop until verif_err_msg = ""
        ButtonPressed = verif_button
    End If

end function

function write_interview_CASE_NOTE()

	' 'Now we case note!
	STATS_manualtime = STATS_manualtime + 600
	Call start_a_blank_case_note

	If create_incomplete_note_checkbox = checked then
		CALL write_variable_in_CASE_NOTE("Partial Interview Information from " & interview_date)
	Else
		CALL write_variable_in_CASE_NOTE("~ Interview Completed on " & interview_date & " ~")
	End If
	If run_by_interview_team = True Then
		CALL write_variable_in_CASE_NOTE("--Interview completed and no processing work done.")
		CALL write_variable_in_CASE_NOTE("--Processing to be completed by a follow up worker.")
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
	If signature_detail = "Accepted Verbally" or second_signature_detail = "Accepted Verbally" Then
		CALL write_variable_in_CASE_NOTE("* * Verbal Signature Accepted during interview for:")
		If signature_detail = "Accepted Verbally" Then CALL write_variable_in_CASE_NOTE("    - MEMB " & signature_person)
		If second_signature_detail = "Accepted Verbally" Then CALL write_variable_in_CASE_NOTE("    - MEMB " & second_signature_person)
		CALL write_variable_in_CASE_NOTE("    Signature accepted on " & verbal_sig_date & " at " & verbal_sig_time & ".")
		CALL write_variable_in_CASE_NOTE("    Resident Phone Number: " & verbal_sig_phone_number)
	End If
	CALL write_variable_in_CASE_NOTE("Interview Programs:")
    memb_snap_checkbox = unchecked
    memb_cash_checkbox = unchecked
    memb_emer_checkbox = unchecked
    memb_none_checkbox = unchecked
    all_hh_memb_progs_match = True
    If HH_MEMB_ARRAY(snap_req_checkbox, 0) = checked Then memb_snap_checkbox = checked
    If HH_MEMB_ARRAY(cash_req_checkbox, 0) = checked Then memb_cash_checkbox = checked
    If HH_MEMB_ARRAY(emer_req_checkbox, 0) = checked Then memb_emer_checkbox = checked
    If HH_MEMB_ARRAY(none_req_checkbox, 0) = checked Then memb_none_checkbox = checked
	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_members) = False Then
            If HH_MEMB_ARRAY(snap_req_checkbox, the_members) <> memb_snap_checkbox Then all_hh_memb_progs_match = False
            If HH_MEMB_ARRAY(cash_req_checkbox, the_members) <> memb_cash_checkbox Then all_hh_memb_progs_match = False
            If HH_MEMB_ARRAY(emer_req_checkbox, the_members) <> memb_emer_checkbox Then all_hh_memb_progs_match = False
            If HH_MEMB_ARRAY(none_req_checkbox, the_members) <> memb_none_checkbox Then all_hh_memb_progs_match = False
        End If
    Next

	If cash_request = True Then
		If the_process_for_cash = "Application" Then CALL write_variable_in_CASE_NOTE(" - CASH at Application. App Date: " & CAF_datestamp & ". " & type_of_cash & " Cash.")
		If the_process_for_cash = "Renewal" Then CALL write_variable_in_CASE_NOTE(" - CASH at Renewal. Renewal Month: " & next_cash_revw_mo & "/" & next_cash_revw_yr & ". " & type_of_cash & " Cash.")
		If cash_verbal_request = "Yes" Then CALL write_variable_in_CASE_NOTE("   -CASH requested verbally during the interview")
		If cash_verbal_withdraw = "Yes" Then CALL write_variable_in_CASE_NOTE("   -VERBAL WITHDRAW OF CASH REQUEST")
	End If
	If grh_request = True Then
		If the_process_for_grh = "Application" Then CALL write_variable_in_CASE_NOTE(" - HS/GRH at Application. App Date: " & CAF_datestamp & ".")
		If the_process_for_grh = "Renewal" Then CALL write_variable_in_CASE_NOTE(" - HS/GRH at Renewal. Renewal Month: " & next_grh_revw_mo & "/" & next_grh_revw_yr & ".")
		If grh_verbal_request = "Yes" 		Then CALL write_variable_in_CASE_NOTE("   -HS/GRH requested verbally during the interview")
		If grh_verbal_withdraw = "Yes" Then CALL write_variable_in_CASE_NOTE("   -VERBAL WITHDRAW OF HS/GRH REQUEST")
	End If
	If snap_request = True Then
		If the_process_for_snap = "Application" Then CALL write_variable_in_CASE_NOTE(" - SNAP at Application. App Date: " & CAF_datestamp & ".")
		If the_process_for_snap = "Renewal" Then CALL write_variable_in_CASE_NOTE(" - SNAP at Renewal. Renewal Month: " & next_snap_revw_mo & "/" & next_snap_revw_yr & ".")
		If snap_verbal_request = "Yes" Then CALL write_variable_in_CASE_NOTE("   -SNAP requested verbally during the interview")
		If snap_verbal_withdraw = "Yes" Then CALL write_variable_in_CASE_NOTE("   -VERBAL WITHDRAW OF SNAP REQUEST")
	End If
	If emer_request = True Then
		CALL write_variable_in_CASE_NOTE(" - EMERGENCY Request at Application. App Date: " & CAF_datestamp & ". EMER is " & type_of_emer)
		If emer_verbal_request = "Yes" Then CALL write_variable_in_CASE_NOTE("   -Emergency requested verbally during the interview")
		If emer_verbal_withdraw = "Yes" Then CALL write_variable_in_CASE_NOTE("   -VERBAL WITHDRAW OF EMERGENCY REQUEST")
	End If
	Call write_bullet_and_variable_in_CASE_NOTE("Program Request Notes", program_request_notes)
	Call write_bullet_and_variable_in_CASE_NOTE("Verbal Request Notes", verbal_request_notes)
    If all_hh_memb_progs_match Then write_variable_in_CASE_NOTE("Program requests include everyone listed on this case.")

    If CAF_form = "MNbenefits" AND (trim(additional_application_comments) <> "" OR trim(additional_income_comments) <> "" OR trim(cover_letter_interview_notes) <> "") Then
        CALL write_variable_in_CASE_NOTE("--- MN Benefits Appliction Cover Letter Details ---")
        If trim(additional_application_comments) <> "" Then CALL write_bullet_and_variable_in_CASE_NOTE("Additional Application Comments", additional_application_comments)
        If trim(additional_income_comments) <> "" Then CALL write_bullet_and_variable_in_CASE_NOTE("Additional Income Comments", additional_income_comments)
        If trim(cover_letter_interview_notes) <> "" Then CALL write_bullet_and_variable_in_CASE_NOTE("Interview Notes on Cover Letter Details", cover_letter_interview_notes)
    End If

	CALL write_variable_in_CASE_NOTE("--- Household Members ---")
    membs_no_request = ""
	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(ignore_person, the_members) = False and HH_MEMB_ARRAY(none_req_checkbox, the_members) = checked Then membs_no_request = membs_no_request & "M" & HH_MEMB_ARRAY(ref_number, the_members) & " - " & HH_MEMB_ARRAY(full_name_const, the_members) & ", "
        If HH_MEMB_ARRAY(ignore_person, the_members) = False and HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked Then
            CALL write_variable_in_CASE_NOTE("  * " & HH_MEMB_ARRAY(ref_number, the_members) & "-" & HH_MEMB_ARRAY(full_name_const, the_members))
            If NOT all_hh_memb_progs_match Then
                prog_list = ""
                If HH_MEMB_ARRAY(snap_req_checkbox, the_members) = checked Then prog_list = prog_list & "SNAP, "
                If HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked Then prog_list = prog_list & "CASH, "
                If HH_MEMB_ARRAY(emer_req_checkbox, the_members) = checked Then prog_list = prog_list & "EMERGENCY, "
                ' If HH_MEMB_ARRAY(none_req_checkbox, the_members) = checked Then prog_list = prog_list & "NO PROGRAMS, "
                If prog_list <> "" Then prog_list = left(prog_list, len(prog_list)-2)
                CALL write_variable_in_CASE_NOTE("    Program Requests: " & prog_list)
            End If
            If HH_MEMB_ARRAY(pers_in_maxis, the_members) = False Then
                CALL write_variable_in_CASE_NOTE("    Person NOT in MAXIS. Details recorded during interview:")
                demographic_details = ""
                If HH_MEMB_ARRAY(date_of_birth, the_members) <> "" Then demographic_details = demographic_details & "DOB: " & HH_MEMB_ARRAY(date_of_birth, the_members) & ", "
                If HH_MEMB_ARRAY(gender, the_members) <> "" Then demographic_details = demographic_details & "Gender: " & HH_MEMB_ARRAY(gender, the_members) & ", "
                If HH_MEMB_ARRAY(rel_to_applcnt, the_members) <> "" Then demographic_details = demographic_details & "Rel to M01: " & HH_MEMB_ARRAY(rel_to_applcnt, the_members) & ", "
                demographic_details = trim(demographic_details)
                If right(demographic_details, 1) = "," Then demographic_details = left(demographic_details, len(demographic_details)-1)
                If demographic_details <> "" Then CALL write_variable_in_CASE_NOTE("    - " & demographic_details)
                If HH_MEMB_ARRAY(ssn, the_members) <> "" Then CALL write_variable_in_CASE_NOTE("    - SSN in Case File on Interview Notes Doc (WIF).")
                If HH_MEMB_ARRAY(id_verif, the_members) <> "" Then          Call write_variable_in_CASE_NOTE("    - ID Verification: " & HH_MEMB_ARRAY(id_verif, the_members) )
                If HH_MEMB_ARRAY(citizen, the_members) <> "" Then           Call write_variable_in_CASE_NOTE("    - Citizen: " & HH_MEMB_ARRAY(citizen, the_members) )
                demographics_plus = ""
                If HH_MEMB_ARRAY(race, the_members) <> "" Then demographics_plus = demographics_plus & "Race: " & HH_MEMB_ARRAY(race, the_members) & ", "
                If HH_MEMB_ARRAY(ethnicity_yn, the_members) <> "" Then demographics_plus = demographics_plus & "Hispanic: " & HH_MEMB_ARRAY(ethnicity_yn, the_members) & ", "
                If HH_MEMB_ARRAY(marital_status, the_members) <> "" Then demographics_plus = demographics_plus & "Marital Status: " & HH_MEMB_ARRAY(marital_status, the_members) & ", "
                If HH_MEMB_ARRAY(last_grade_completed, the_members) <> "" Then demographics_plus = demographics_plus & "Last Grade: " & HH_MEMB_ARRAY(last_grade_completed, the_members) & ", "
                demographics_plus = trim(demographics_plus)
                If right(demographics_plus, 1) = "," Then demographics_plus = left(demographics_plus, len(demographics_plus)-1)
                If demographics_plus <> "" Then CALL write_variable_in_CASE_NOTE("    - " & demographics_plus)
                move_info = ""
                If HH_MEMB_ARRAY(mn_entry_date, the_members) <> "" Then move_info = move_info & "MN Entry Date: " & HH_MEMB_ARRAY(mn_entry_date, the_members) & ", "
                If HH_MEMB_ARRAY(former_state, the_members) <> "" Then move_info = move_info & "Former State: " & HH_MEMB_ARRAY(former_state, the_members) & ", "
                move_info = trim(move_info)
                If right(move_info, 1) = "," Then move_info = left(move_info, len(move_info)-1)
                If move_info <> "" Then CALL write_variable_in_CASE_NOTE("    - " & move_info)

            Else
                If CHANGES_ARRAY(last_name_const, the_members) <> ""      Then Call write_variable_in_CASE_NOTE("    - Last Name Changed from " & CHANGES_ARRAY(last_name_const, the_members) & " to " & HH_MEMB_ARRAY(last_name_const, the_members) )
                If CHANGES_ARRAY(first_name_const, the_members) <> ""     Then Call write_variable_in_CASE_NOTE("    - First Name Changed from " & CHANGES_ARRAY(first_name_const, the_members) & " to " & HH_MEMB_ARRAY(first_name_const, the_members) )
                If CHANGES_ARRAY(mid_initial, the_members) <> ""          Then Call write_variable_in_CASE_NOTE("    - Mid Initial Changed from " & CHANGES_ARRAY(mid_initial, the_members) & " to " & HH_MEMB_ARRAY(mid_initial, the_members) )
                If CHANGES_ARRAY(date_of_birth, the_members) <> ""        Then Call write_variable_in_CASE_NOTE("    - Date of Birth Changed from " & CHANGES_ARRAY(date_of_birth, the_members) & " to " & HH_MEMB_ARRAY(date_of_birth, the_members) )
                If CHANGES_ARRAY(birthdate_verif, the_members) <> ""      Then Call write_variable_in_CASE_NOTE("    - DoB Verif Changed from " & CHANGES_ARRAY(birthdate_verif, the_members) & " to " & HH_MEMB_ARRAY(birthdate_verif, the_members) )
                If CHANGES_ARRAY(age, the_members) <> ""                  Then Call write_variable_in_CASE_NOTE("    - Age Changed from " & CHANGES_ARRAY(age, the_members) & " to " & HH_MEMB_ARRAY(age, the_members) )
                If CHANGES_ARRAY(ssn, the_members) <> ""                  Then Call write_variable_in_CASE_NOTE("    - SSN Updated.")'" from " & CHANGES_ARRAY(ssn, the_members) & " to " & HH_MEMB_ARRAY(ssn, the_members) )
                If CHANGES_ARRAY(ssn_verif, the_members) <> ""            Then Call write_variable_in_CASE_NOTE("    - SSN Verif Changed from " & CHANGES_ARRAY(ssn_verif, the_members) & " to " & HH_MEMB_ARRAY(ssn_verif, the_members) )
                If CHANGES_ARRAY(spoken_lang, the_members) <> ""          Then Call write_variable_in_CASE_NOTE("    - Spoken Lang Changed from " & CHANGES_ARRAY(spoken_lang, the_members) & " to " & HH_MEMB_ARRAY(spoken_lang, the_members) )
                If CHANGES_ARRAY(written_lang, the_members) <> ""         Then Call write_variable_in_CASE_NOTE("    - Written Lang Changed from " & CHANGES_ARRAY(written_lang, the_members) & " to " & HH_MEMB_ARRAY(written_lang, the_members) )
                If CHANGES_ARRAY(interpreter, the_members) <> ""          Then Call write_variable_in_CASE_NOTE("    - Interpreter Needed Changed from " & CHANGES_ARRAY(interpreter, the_members) & " to " & HH_MEMB_ARRAY(interpreter, the_members) )
                If CHANGES_ARRAY(alias_yn, the_members) = "Y"             Then Call write_variable_in_CASE_NOTE("    - Alias Name Added: " & HH_MEMB_ARRAY(other_names, the_members) )
                If CHANGES_ARRAY(alias_yn, the_members) = "N"             Then Call write_variable_in_CASE_NOTE("    - Alias Name Removed." )
                ' If CHANGES_ARRAY(alias_yn, the_members) <> ""           Then Call write_variable_in_CASE_NOTE("    - XXXX Changed from " & CHANGES_ARRAY(alias_yn, the_members) & " to " & HH_MEMB_ARRAY(alias_yn, the_members) )
                If CHANGES_ARRAY(gender, the_members) <> ""               Then Call write_variable_in_CASE_NOTE("    - Gender Changed from " & CHANGES_ARRAY(gender, the_members) & " to " & HH_MEMB_ARRAY(gender, the_members) )
                If CHANGES_ARRAY(race, the_members) <> ""                 Then Call write_variable_in_CASE_NOTE("    - Race Changed from " & CHANGES_ARRAY(race, the_members) & " to " & HH_MEMB_ARRAY(race, the_members) )
                If CHANGES_ARRAY(ethnicity_yn, the_members) <> ""         Then Call write_variable_in_CASE_NOTE("    - Ethnicity Changed from " & CHANGES_ARRAY(ethnicity_yn, the_members) & " to " & HH_MEMB_ARRAY(ethnicity_yn, the_members) )
                If CHANGES_ARRAY(rel_to_applcnt, the_members) <> ""       Then Call write_variable_in_CASE_NOTE("    - Rel to Applicant Changed from " & CHANGES_ARRAY(rel_to_applcnt, the_members) & " to " & HH_MEMB_ARRAY(rel_to_applcnt, the_members) )
                If CHANGES_ARRAY(id_verif, the_members) <> ""             Then Call write_variable_in_CASE_NOTE("    - ID Verif Changed from " & CHANGES_ARRAY(id_verif, the_members) & " to " & HH_MEMB_ARRAY(id_verif, the_members) )
                If CHANGES_ARRAY(marital_status, the_members) <> ""       Then Call write_variable_in_CASE_NOTE("    - Marital Status Changed from " & CHANGES_ARRAY(marital_status, the_members) & " to " & HH_MEMB_ARRAY(marital_status, the_members) )
                If CHANGES_ARRAY(last_grade_completed, the_members) <> "" Then Call write_variable_in_CASE_NOTE("    - Last Grade Changed from " & CHANGES_ARRAY(last_grade_completed, the_members) & " to " & HH_MEMB_ARRAY(last_grade_completed, the_members) )
                If CHANGES_ARRAY(citizen, the_members) <> ""              Then Call write_variable_in_CASE_NOTE("    - Citizen Changed from " & CHANGES_ARRAY(citizen, the_members) & " to " & HH_MEMB_ARRAY(citizen, the_members) )
                If CHANGES_ARRAY(mn_entry_date, the_members) <> ""        Then Call write_variable_in_CASE_NOTE("    - MN Entry Date Changed from " & CHANGES_ARRAY(mn_entry_date, the_members) & " to " & HH_MEMB_ARRAY(mn_entry_date, the_members) )
                If CHANGES_ARRAY(former_state, the_members) <> ""         Then Call write_variable_in_CASE_NOTE("    - Former State Changed from " & CHANGES_ARRAY(former_state, the_members) & " to " & HH_MEMB_ARRAY(former_state, the_members) )

            End If

            If the_members = 0 Then CALL write_variable_in_CASE_NOTE("    Identity: " & HH_MEMB_ARRAY(id_verif, the_members))
            If HH_MEMB_ARRAY(citizen, the_members) = "No" Then
                CALL write_variable_in_CASE_NOTE("    Member is a Non-Citizen.")
                If HH_MEMB_ARRAY(clt_has_sponsor, the_members) <> "?" Then CALL write_variable_in_CASE_NOTE("    * Sponsor: " & HH_MEMB_ARRAY(clt_has_sponsor, the_members))
                If trim(HH_MEMB_ARRAY(imig_status, the_members)) <> "" Then CALL write_variable_in_CASE_NOTE("    * IMIG NOTES: " & HH_MEMB_ARRAY(imig_status, the_members))
            End If
    		If trim(HH_MEMB_ARRAY(client_notes, the_members)) <> "" Then CALL write_variable_in_CASE_NOTE("    NOTES: " & HH_MEMB_ARRAY(client_notes, the_members))
            If HH_MEMB_ARRAY(client_verification, the_members) <> "Not Needed" and HH_MEMB_ARRAY(client_verification, the_members) <> "" Then
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
    membs_no_request = trim(membs_no_request)
    If membs_no_request <> "" Then
        If right(membs_no_request, 1) = "," Then membs_no_request = left(membs_no_request, len(membs_no_request)-1)
        CALL write_variable_in_CASE_NOTE("MEMBS NO Request: " & membs_no_request)
    End If

    If all_members_listed_yn <> "" Then CALL write_variable_in_CASE_NOTE("* ALL HH members Listed: " & all_members_listed_yn)
    If trim(all_members_listed_notes) <> "" Then CALL write_variable_in_CASE_NOTE("* HH Comp Notes: " & all_members_listed_notes)

    If all_members_in_MN_yn <> "" Then
        CALL write_variable_in_CASE_NOTE("* ALL HH members Intend to reside in MN: " & all_members_in_MN_yn)
        If trim(all_members_in_MN_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Notes: " & all_members_in_MN_notes)
    End If
    If anyone_pregnant_yn <> "" Then
        CALL write_variable_in_CASE_NOTE("* Anyone Pregnant: " & anyone_pregnant_yn)
        If trim(anyone_pregnant_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Notes: " & anyone_pregnant_notes)
    End If
    If anyone_served_yn <> "" Then
        CALL write_variable_in_CASE_NOTE("* Anyone Served in Military: " & anyone_served_yn)
        If trim(anyone_served_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Notes: " & anyone_served_notes)
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
	CALL write_variable_in_CASE_NOTE("Housing Support Information: ")
	If trim(licensed_facility) <> "" Then CALL write_variable_in_CASE_NOTE("    Currently reside in licensed facility? " & licensed_facility)
	If trim(meal_provided) <> "" Then CALL write_variable_in_CASE_NOTE("    Residence provides meals? " & meal_provided)
	If trim(residence_name_phone) <> "" Then CALL write_variable_in_CASE_NOTE("    Name/phone number of residence: " & residence_name_phone)
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
	If send_text = "Yes" Then call write_variable_in_CASE_NOTE("  - Send Updates Via Text Message: Yes")
	If send_text = "No" Then call write_variable_in_CASE_NOTE("  - Send Updates Via Text Message: No")
	If send_email = "Yes" Then call write_variable_in_CASE_NOTE("  - Send Updates Via E-Mail: Yes")
	If send_email = "No" Then call write_variable_in_CASE_NOTE("  - Send Updates Via E-Mail: No")

	CALL write_variable_in_CASE_NOTE("-----  CAF Information and Notes -----")

	If trim(pwe_selection) <> "" AND pwe_selection <> "Select or Type" Then CALL write_variable_in_CASE_NOTE("PWE: " & pwe_selection)

	For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
		FORM_QUESTION_ARRAY(each_question).enter_case_note()
	Next

	If disc_rent_amounts = "RESOLVED" or disc_yes_phone_no_expense = "RESOLVED" or disc_no_phone_yes_expense = "RESOLVED" or disc_utility_amounts = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("-----  Answer Clarifications  -----")
	End If
	If disc_rent_amounts = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    HOUSING EXPENSE FORM DETAILS MAY NOT MATCH IN ALL QUESTIONS")
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

	If disc_utility_amounts = "RESOLVED" Then
		CALL write_variable_in_CASE_NOTE("    UTILITY EXPENSE FORM DETAILS NOT MATCH IN ALL QUESTIONS")
		CALL write_variable_in_CASE_NOTE("    Resolution: " & disc_utility_amounts_confirmation)
	End If


	'If at least one field is filled in, then it will write the emergency info to case note
	If resident_emergency_yn <> " " or (trim(emergency_type) <> "" and trim(emergency_type) <> "Select or Type") or trim(emergency_discussion) <> "" or trim(emergency_amount) <> "" or trim(emergency_deadline) <> "" Then
		CALL write_variable_in_CASE_NOTE("-----  Emergency Questions -----")
		If resident_emergency_yn <> " " Then CALL write_variable_in_CASE_NOTE("      Resident experiencing an emergency - " & resident_emergency_yn)
		If trim(emergency_type) <> "" and trim(emergency_type) <> "Select or Type" Then CALL write_variable_in_CASE_NOTE("      Type of emergency - " & emergency_type)
		If trim(emergency_discussion) <> "" Then CALL write_variable_in_CASE_NOTE("      Discussion of emergency with resident - " & emergency_discussion)
		If trim(emergency_amount) <> ""  Then CALL write_variable_in_CASE_NOTE("      Amount needed to resolve emergency - " & emergency_amount)
		If trim(emergency_deadline) <> "" Then CALL write_variable_in_CASE_NOTE("      Deadline to resolve emergency - " & emergency_deadline)
	End If

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
	If trim(edrs_notes_for_case) <> "" Then CALL write_variable_in_CASE_NOTE("      EDRS Notes - " & edrs_notes_for_case)

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

	If run_by_interview_team = True Then
		Call write_variable_in_CASE_NOTE("Read Rights and Responsibilites to resident.")
		If snap_case = True Then
			Call write_variable_in_CASE_NOTE("SNAP E&T Assistance:")
			Call write_bullet_and_variable_in_CASE_NOTE("Q1: Anyone interested in E&T assistance now?", interested_in_job_assistance_now)
			Call write_bullet_and_variable_in_CASE_NOTE("         HH Memb Interested: ", interested_names_now)
			Call write_bullet_and_variable_in_CASE_NOTE("Q2: Anyone interested in E&T assistance in the future?", interested_in_job_assistance_future)
			Call write_bullet_and_variable_in_CASE_NOTE("         HH Memb Interested: ", interested_names_future)
		End If
	Else
		'R&R
		forms_reviewed = ""
		If DHS_4163_checkbox = checked Then forms_reviewed = forms_reviewed & " -4163 -EBT Info"
		If DHS_3315A_checkbox = checked Then forms_reviewed = forms_reviewed & " -3315A"
		If DHS_3979_checkbox = checked Then forms_reviewed = forms_reviewed & " -3979"
		If DHS_2759_checkbox = checked Then forms_reviewed = forms_reviewed & " -2759"
		If DHS_3353_checkbox = checked Then forms_reviewed = forms_reviewed & " -3353"
		If DHS_2920_checkbox = checked Then forms_reviewed = forms_reviewed & " -2920"
		If DHS_3477_checkbox = checked Then forms_reviewed = forms_reviewed & " -3477"
		If DHS_4133_checkbox = checked Then forms_reviewed = forms_reviewed & " -4133"
		If DHS_2647_checkbox = checked Then forms_reviewed = forms_reviewed & " -2647"
		If DHS_2929_checkbox = checked Then forms_reviewed = forms_reviewed & " -2929"
		If DHS_3323_checkbox = checked Then forms_reviewed = forms_reviewed & " -3323"
		If DHS_3393_checkbox = checked Then forms_reviewed = forms_reviewed & " -3393"
		If DHS_3163B_checkbox = checked Then forms_reviewed = forms_reviewed & " -3163B"
		If DHS_2338_checkbox = checked Then forms_reviewed = forms_reviewed & " -2338"
		If DHS_5561_checkbox = checked Then forms_reviewed = forms_reviewed & " -5561"
		If DHS_2961_checkbox = checked Then forms_reviewed = forms_reviewed & " -2961"
		If DHS_2887_checkbox = checked Then forms_reviewed = forms_reviewed & " -2887"
		If DHS_3238_checkbox = checked Then forms_reviewed = forms_reviewed & " -3238"
		If DHS_2625_checkbox = checked Then forms_reviewed = forms_reviewed & " -2625 -7635"

		If left(forms_reviewed, 2) = " -" Then forms_reviewed = right(forms_reviewed, len(forms_reviewed)-2)
		Call write_bullet_and_variable_in_CASE_NOTE("Reviewed DHS Forms", forms_reviewed)
		If DHS_2625_checkbox = checked Then
			Call write_variable_in_CASE_NOTE("SNAP Reporting discussed. Case appears to be a " & snap_reporting_type & " reporter.")
			Call write_variable_in_CASE_NOTE("     Next review month of " & next_revw_month)
			Call write_variable_in_CASE_NOTE("     This may change dependent on info received up until SNAP approval.")
		End If
	End If

	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	If run_by_interview_team = True Then
		PF3
		Call start_a_blank_CASE_NOTE
		Call write_variable_in_CASE_NOTE("Processing Needed: follow up notes from Interview")
		Call write_variable_in_CASE_NOTE("Interview completed on " & interview_date & " with " & who_are_we_completing_the_interview_with & ".")
		If is_elig_XFS = True Then
			Call write_variable_in_CASE_NOTE("*** SNAP APPEARS TO NEED EXPEDITED PROCESSING ***")
			Call write_variable_in_case_note("  Based on: Income:  $ " & right("        " & determined_income, 8) & ", Assets:    $ " & right("        " & determined_assets, 8)   & ", Totaling: $ " & right("        " & calculated_resources, 8))
			Call write_variable_in_case_note("            Shelter: $ " & right("        " & determined_shel, 8)   & ", Utilities: $ " & right("        " & determined_utilities, 8) & ", Totaling: $ " & right("        " & calculated_expenses, 8))
		End If
		Call write_variable_in_CASE_NOTE("---")

		caf_progs = ""
		If CASH_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", Cash"
		If GRH_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", HS/GRH"
		If SNAP_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", SNAP"
		If EMER_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", EMER"
		If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
		If caf_progs <> "" Then CALL write_variable_in_CASE_NOTE("Programs Requested on CAF: " & caf_progs)

		progs_verbal_request = ""
		If cash_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", Cash"
		If grh_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", HS/GRH"
		If snap_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", SNAP"
		If emer_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", EMER"
		If left(progs_verbal_request, 2) = ", " Then progs_verbal_request = right(progs_verbal_request, len(progs_verbal_request)-2)
		If progs_verbal_request <> "" Then CALL write_variable_in_CASE_NOTE("Programs Requested Verbally: " & progs_verbal_request)

		progs_withdraw_request = ""
		If cash_verbal_withdraw = "Yes" Then progs_withdraw_request = progs_withdraw_request & ", Cash"
		If grh_verbal_withdraw = "Yes" Then progs_withdraw_request = progs_withdraw_request & ", HS/GRH"
		If snap_verbal_withdraw = "Yes" Then progs_withdraw_request = progs_withdraw_request & ", SNAP"
		If emer_verbal_withdraw = "Yes" Then progs_withdraw_request = progs_withdraw_request & ", EMER"
		If left(progs_withdraw_request, 2) = ", " Then progs_withdraw_request = right(progs_withdraw_request, len(progs_withdraw_request)-2)
		If progs_withdraw_request <> "" Then
			Call write_variable_in_CASE_NOTE("* * * Resident Requested Withdraw of Programs * * *")
			CALL write_variable_in_CASE_NOTE("      Programs Requested Withdraw: " & progs_withdraw_request)
		End If
		Call write_bullet_and_variable_in_CASE_NOTE("Program Request Notes", program_request_notes)
		Call write_bullet_and_variable_in_CASE_NOTE("Verbal Request Notes", verbal_request_notes)
		Call write_variable_in_CASE_NOTE("---")

		Call write_variable_in_CASE_NOTE("For follow up questions or information, resident can be reached at:")
		Call write_variable_in_CASE_NOTE("  --- " & phone_number_selection)
		Call write_variable_in_CASE_NOTE("  --- Leave a detailed message at this number: " & leave_a_message)
		If resident_questions <> "" Then
			Call write_variable_in_CASE_NOTE("Questions/Requests from the Resident:")
			Call write_variable_in_CASE_NOTE("- " & resident_questions)
			Call write_variable_in_CASE_NOTE("---")
		End If
		If case_summary <> "" Then
			Call write_variable_in_CASE_NOTE("Notes from the interviewer:")
			Call write_variable_in_CASE_NOTE("- " & case_summary)
			Call write_variable_in_CASE_NOTE("---")
		End If
		Call write_variable_in_CASE_NOTE("Program History")
		history_found = False
		If snap_closed_in_past_30_days = True or snap_closed_in_past_4_months = True Then
			Call write_bullet_and_variable_in_CASE_NOTE("SNAP recently closed", FS_date_closed & " - " & FS_reason_closed)
			history_found = True
		End If
		If grh_closed_in_past_30_days = True or grh_closed_in_past_4_months = True Then
			Call write_bullet_and_variable_in_CASE_NOTE("GRH recently closed", GRH_date_closed & " - " & GRH_reason_closed)
			history_found = True
		End If
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
			Call write_bullet_and_variable_in_CASE_NOTE(cash1_recently_closed_program & " recently closed", cash1_date_closed & " - " & cash1_closed_reason)
			history_found = True
		End If
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
			Call write_bullet_and_variable_in_CASE_NOTE(cash2_recently_closed_program & " recently closed", cash2_date_closed & " - " & cash2_closed_reason)
			history_found = True
		End If
		If issued_date <> "" Then
			Call write_bullet_and_variable_in_CASE_NOTE("EMER last issued", issued_date & " (" & issued_prog & ")")
			history_found = True
		End If
		If history_found = False Then
			Call write_variable_in_CASE_NOTE("* No recent Program History listed in MX that appears relevant.")
		End If
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)
		PF3
	End If

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

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If FORM_QUESTION_ARRAY(quest).verif_status = "Requested" Then
			verifs_needed = verifs_needed & "; " & FORM_QUESTION_ARRAY(quest).verif_verbiage
			If FORM_QUESTION_ARRAY(quest).verif_notes <> "" Then verifs_needed = verifs_needed & " - " & FORM_QUESTION_ARRAY(quest).verif_notes
		End If
		If FORM_QUESTION_ARRAY(quest).detail_array_exists = True Then
			For each_item = 0 to UBound(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				If FORM_QUESTION_ARRAY(quest).detail_verif_status(each_item) = "Requested" Then
					item_information = ""
					If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " at " & FORM_QUESTION_ARRAY(quest).detail_business(each_item)
					If FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " of " & FORM_QUESTION_ARRAY(quest).detail_type(each_item)
					If FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " of " & FORM_QUESTION_ARRAY(quest).detail_type(each_item)
					If FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then item_information = FORM_QUESTION_ARRAY(quest).detail_type(each_item)
					If FORM_QUESTION_ARRAY(quest).detail_source = "expense" Then item_information = "expense paid by " & FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item)
					If FORM_QUESTION_ARRAY(quest).detail_source = "changes" Then item_information = "change"
					If FORM_QUESTION_ARRAY(quest).detail_source = "winnings" Then item_information = FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) & " winnings"

					verifs_needed = verifs_needed & "; " & FORM_QUESTION_ARRAY(quest).verif_verbiage & " - " & item_information
					If FORM_QUESTION_ARRAY(quest).detail_verif_notes(each_item) <> "" Then verifs_needed = verifs_needed & " - " & FORM_QUESTION_ARRAY(quest).detail_verif_notes(each_item)
				End If
			Next
		End If
	Next

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

			'TODO - Add information from the form answers

			dlg_len = 45 + jobs_grp_len + busi_grp_len + unea_grp_len
			If interview_note_details_exists = True Then dlg_len = dlg_len + intvw_notes_len + 10

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, dlg_len, "Determination of Income in Month of Application"
			  	ButtonGroup ButtonPressed
					'displaying details from interview notes in the dialog for calculating app month income
				  	y_pos = 10
					If interview_note_details_exists = True Then
						GroupBox 10, y_pos, 380, intvw_notes_len, "Interview NOTES entered into the Script already"
						y_pos = y_pos + 15
						' If trim(question_8_interview_notes) <> "" Then
						' 	Text 20, y_pos, 360, 10, "8. Has anyone in the household had a job or been self-employed?"
						' 	Text 30, y_pos+10, 350, 10, question_8_interview_notes
						' 	y_pos = y_pos + 20
						' End If

						y_pos = y_pos + 10

					End If
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

	If run_by_interview_team = True Then
		determined_utilities = 0
		If heat_exp_checkbox = checked OR ac_exp_checkbox = checked Then
			determined_utilities = determined_utilities + heat_AC_amt
		Else
			If electric_exp_checkbox = checked Then determined_utilities = determined_utilities + electric_amt
			If phone_exp_checkbox = checked Then determined_utilities = determined_utilities + phone_amt
		End If
	End If

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

	If run_by_interview_team = True Then
		If is_elig_XFS = True Then case_assesment_text = "Case IS EXPEDITED"
		If is_elig_XFS = False Then case_assesment_text = "Case is NOT EXPEDITED"
		next_steps_one = "Income - $ " & determined_income & ", Assets - $ " & determined_assets & ", Housing - $ " & determined_shel & ", Utilities - $ " & determined_utilities
	End If

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
	Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", True)
	'Call create_outlook_email("", "HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", False)
	'create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
end function
'---------------------------------------------------------------------------------------------------------------------------


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Const end_of_doc = 6			'This is for word document ennumeration

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
' worker_name = user_ID_for_validation
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)

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
memb_panel_relationship_list = memb_panel_relationship_list+chr(9)+"01 Self"
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
marital_status_list = marital_status_list+chr(9)+"L  Legally Seperated"
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

ssn_verif_list = ""
ssn_verif_list = ssn_verif_list+chr(9)+"A - SSN Applied For"
ssn_verif_list = ssn_verif_list+chr(9)+"P - SSN Provided, verif Pending"
ssn_verif_list = ssn_verif_list+chr(9)+"N - SSN Not Provided"
ssn_verif_list = ssn_verif_list+chr(9)+"N - Member Does Not Have SSN"
ssn_verif_list = ssn_verif_list+chr(9)+"V - SSN Verified via Interface"

question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"

Set wshshell = CreateObject("WScript.Shell")						'creating the wscript method to interact with the system
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"	'defining the my documents folder for use in saving script details/variables between script runs

'Dimming all the variables because they are defined and set within functions
Dim HH_arrived_date, HH_arrived_place
Dim who_are_we_completing_the_interview_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, exp_q_4_utilities_this_month, caf_exp_pay_heat_checkbox, caf_exp_pay_ac_checkbox, caf_exp_pay_electricity_checkbox, caf_exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim licensed_facility, meal_provided, residence_name_phone
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_one_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county, CAF_datestamp, all_the_clients, err_msg, interpreter_information, interpreter_language, arep_interview_id_information, non_applicant_interview_info
Dim send_text, send_email, ssn_update_success, ssn_update_attempt
Dim all_members_listed_yn, all_members_listed_notes, all_members_in_MN_yn, all_members_in_MN_notes, anyone_pregnant_yn, anyone_pregnant_notes, anyone_served_yn, anyone_served_notes

Dim intv_app_month_income, intv_app_month_asset, intv_app_month_housing_expense, intv_exp_pay_heat_checkbox, intv_exp_pay_ac_checkbox, intv_exp_pay_electricity_checkbox, intv_exp_pay_phone_checkbox, intv_exp_pay_none_checkbox
Dim id_verif_on_file, snap_active_in_other_state, last_snap_was_exp, how_are_we_completing_the_interview
Dim cash_other_req_detail, snap_other_req_detail, emer_other_req_detail, family_cash_program, famliy_cash_notes

Dim CASH_ever_active, MSA_ever_active, FS_ever_active, MA_ever_active, EMER_ever_active, GRH_ever_active, GA_ever_active, MFIP_ever_active, DWP_ever_active
Dim QMB_ever_active, SLMB_ever_active, CCAP_ever_active, QI1_ever_active, RCA_ever_active, IV_E_ever_active, IMD_ever_active
Dim CASH_currently_active, MSA_currently_active, FS_currently_active, MA_currently_active, EMER_currently_active, GRH_currently_active, GA_currently_active, MFIP_currently_active, DWP_currently_active
Dim QMB_currently_active, SLMB_currently_active, CCAP_currently_active, QI1_currently_active, RCA_currently_active, IV_E_currently_active, IMD_currently_active
Dim CASH_date_closed, MSA_date_closed, FS_date_closed, MA_date_closed, EMER_date_closed, GRH_date_closed, GA_date_closed, MFIP_date_closed, DWP_date_closed
Dim QMB_date_closed, SLMB_date_closed, CCAP_date_closed, QI1_date_closed, RCA_date_closed, IV_E_date_closed, IMD_date_closed
Dim CASH_reason_closed, MSA_reason_closed, FS_reason_closed, MA_reason_closed, EMER_reason_closed, GRH_reason_closed, GA_reason_closed, MFIP_reason_closed, DWP_reason_closed
Dim QMB_reason_closed, SLMB_reason_closed, CCAP_reason_closed, QI1_reason_closed, RCA_reason_closed, IV_E_reason_closed, IMD_reason_closed, active_spans_array
Dim snap_closed_in_past_30_days, snap_closed_in_past_4_months, grh_closed_in_past_30_days, grh_closed_in_past_4_months, issued_date, issued_prog
Dim cash1_closed_in_past_30_days, cash1_closed_in_past_4_months, cash1_recently_closed_program, cash1_date_closed, cash1_closed_reason
Dim cash2_closed_in_past_30_days, cash2_closed_in_past_4_months, cash2_recently_closed_program, cash2_date_closed, cash2_closed_reason

Dim additional_application_comments, additional_income_comments, cover_letter_interview_notes
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_three, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five
Dim arep_name, arep_relationship, arep_phone_number, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, need_to_update_addr
Dim MAXIS_arep_name, MAXIS_arep_relationship, MAXIS_arep_phone_number, MAXIS_arep_addr_street, MAXIS_arep_addr_city, MAXIS_arep_addr_state, MAXIS_arep_addr_zip
Dim CAF_arep_name, CAF_arep_relationship, CAF_arep_phone_number, CAF_arep_addr_street, CAF_arep_addr_city, CAF_arep_addr_state, CAF_arep_addr_zip
Dim arep_complete_forms_checkbox, arep_get_notices_checkbox, arep_use_SNAP_checkbox
Dim CAF_arep_complete_forms_checkbox, CAF_arep_get_notices_checkbox, CAF_arep_use_SNAP_checkbox
Dim arep_on_CAF_checkbox, arep_action, CAF_arep_action, arep_and_CAF_arep_match, arep_authorization, arep_exists, arep_authorized
Dim signature_detail, signature_person, second_signature_detail, second_signature_person
Dim client_signed_verbally_yn, interview_date, add_to_time, update_arep, verifs_needed, verifs_selected, verif_req_form_sent_date, number_verifs_checkbox, verifs_postponed_checkbox
Dim verif_snap_checkbox, verif_cash_checkbox, verif_mfip_checkbox, verif_dwp_checkbox, verif_msa_checkbox, verif_ga_checkbox, verif_grh_checkbox, verif_emer_checkbox, verif_hc_checkbox
Dim exp_snap_approval_date, exp_snap_delays, snap_denial_date, snap_denial_explain, pend_snap_on_case, do_we_have_applicant_id
Dim resident_emergency_yn, emergency_type, emergency_discussion, emergency_amount, emergency_deadline
Dim family_cash_case_yn, absent_parent_yn, relative_caregiver_yn, minor_caregiver_yn
Dim pwe_selection
Dim disc_phone_confirmation, disc_yes_phone_no_expense_confirmation, disc_no_phone_yes_expense_confirmation, disc_homeless_confirmation, disc_out_of_county_confirmation, CAF1_rent_indicated, Verbal_rent_indicated
Dim Q14_rent_indicated, rent_summary, disc_rent_amounts_confirmation, disc_utility_caf_1_summary, utility_summary, disc_utility_amounts_confirmation
Dim qual_numb, exp_num, last_num, emer_numb, discrep_num, verbal_sig_date, verbal_sig_time, verbal_sig_phone_number

'R&R
Dim DHS_4163_checkbox, DHS_3315A_checkbox, DHS_3979_checkbox, DHS_2759_checkbox, DHS_3353_checkbox, DHS_2920_checkbox, DHS_3477_checkbox, DHS_4133_checkbox, DHS_2647_checkbox
Dim DHS_2929_checkbox, DHS_3323_checkbox, DHS_3393_checkbox, DHS_3163B_checkbox, DHS_2338_checkbox, DHS_5561_checkbox, DHS_2961_checkbox, DHS_2887_checkbox, DHS_3238_checkbox, DHS_2625_checkbox
Dim case_card_info, clt_knows_how_to_use_ebt_card, snap_reporting_type, next_revw_month, confirm_recap_read, confirm_cover_letter_read, case_summary, phone_number_selection, leave_a_message, resident_questions
Dim interested_in_job_assistance_now, interested_names_now, interested_in_job_assistance_future, interested_names_future
Dim cash_request, snap_request, emer_request, grh_request
Dim show_pg_one_memb01_and_exp, show_pg_one_address, show_pg_memb_list, show_q_1_6
Dim show_q_7_11, show_q_14_15, show_q_21_24, show_qual, show_pg_last, discrepancy_questions, show_arep_page, expedited_determination, emergency_questions
Dim CASH_on_CAF_checkbox, SNAP_on_CAF_checkbox, EMER_on_CAF_checkbox, GRH_on_CAF_checkbox, verbal_request_notes, program_request_notes
Dim cash_verbal_request, snap_verbal_request, emer_verbal_request, grh_verbal_request, cash_verbal_withdraw, snap_verbal_withdraw, emer_verbal_withdraw, grh_verbal_withdraw
Dim type_of_cash, the_process_for_cash, next_cash_revw_mo, next_cash_revw_yr
Dim the_process_for_snap, next_snap_revw_mo, next_snap_revw_yr
Dim the_process_for_grh, next_grh_revw_mo, next_grh_revw_yr
Dim type_of_emer, the_process_for_emer


'EXPEDITED DETERMINATION VARIABLES'
Dim exp_det_income, exp_det_assets, exp_det_housing, exp_det_utilities, heat_exp_checkbox, ac_exp_checkbox, electric_exp_checkbox, phone_exp_checkbox, none_exp_checkbox, exp_det_notes
Dim expedited_determination_completed, determined_income, determined_assets, determined_shel, determined_utilities, calculated_resources
Dim jobs_income_yn, busi_income_yn, unea_income_yn, cash_amount_yn, bank_account_yn, all_utilities, heat_expense, ac_expense, electric_expense, phone_expense, none_expense, expedited_screening
Dim calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, case_is_expedited, approval_date, caf_1_resources, caf_1_expenses
' Dim calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS, approval_date, CAF_datestamp, interview_date
Dim applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four
' Dim applicant_id_on_file_yn, applicant_id_through_SOLQ, delay_explanation, snap_denial_date, snap_denial_explain, case_assesment_text, next_steps_one, next_steps_two, next_steps_three, next_steps_four
Dim postponed_verifs_yn, list_postponed_verifs, day_30_from_application, other_snap_state, other_state_reported_benefit_end_date, other_state_benefits_openended, other_state_contact_yn
Dim other_state_verified_benefit_end_date, mn_elig_begin_date, action_due_to_out_of_state_benefits, case_has_previously_postponed_verifs_that_prevent_exp_snap, prev_post_verif_assessment_done
Dim rent_amount, lot_rent_amount, mortgage_amount, insurance_amount, tax_amount, room_amount, garage_amount, cash_amount
Dim previous_CAF_datestamp, previous_expedited_package, prev_verifs_mandatory_yn, prev_verif_list, curr_verifs_postponed_yn, ongoing_snap_approved_yn, prev_post_verifs_recvd_yn
Dim delay_action_due_to_faci, deny_snap_due_to_faci, faci_review_completed, facility_name, snap_inelig_faci_yn, faci_entry_date, faci_release_date, release_date_unknown_checkbox, release_within_30_days_yn
Dim income_review_completed, assets_review_completed, shel_review_completed, note_calculation_detail

ssn_update_attempt = False
ssn_update_success = False

show_cover_letter = 0
show_pg_one_memb01_and_exp	= 1
show_pg_one_address			= 2
show_pg_memb_list			= 3
' show_q_1_6					= 4
' show_q_7_11					= 5
' show_q_12_13				= 6
' show_q_14_15				= 7
' show_q_16_20				= 8
' show_q_21_24				= 9
show_qual					= 10
show_pg_last				= 11
discrepancy_questions		= 12
show_arep_page				= 20
expedited_determination		= 14
emergency_questions 		= 15


show_exp_pg_amounts = 1
show_exp_pg_determination = 2
show_exp_pg_review = 3

update_addr = FALSE
update_pers = FALSE
need_to_update_addr = FALSE
page_display = 1
discrepancies_exist = False
children_under_18_in_hh = False
children_under_22_in_hh = False
school_age_children_in_hh = False
expedited_determination_needed = False
expedited_determination_completed = False
first_time_in_exp_det = True
expedited_viewed = False

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
	If continue_in_inquiry = vbNo Then
		STATS_manualtime = STATS_manualtime + (timer - start_time)
		Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
	End If
End If
If MX_region = "TRAINING" Then developer_mode = True


'look to see if the worker is listed as one of the interviewer workers
run_by_interview_team = False										'Default the interview team option to false
For each worker in interviewer_array 								'loop through all of the workers listed in the interviewer_array
	If user_ID_for_validation = worker.interviewer_id_number Then					'if the worker county logon ID that is running the script matches one of the interviewer_array workers
		run_interview_team_msg = ""
		If worker.interview_trainer = True Then
			run_interview_team_msg = MsgBox("The Interview Script has two run options." & vbCr & vbCr & "- One is for a standard worker that has full policy and processing knowledge/training." & vbCr & "- The other is for the team of workers that complete interviews only and no processing." & vbCr & vbCr & "Do you want to run the Interview Team - INTERVIEW ONLY NO PROCESSING - Option?", vbQuestion + vbYesNo, "Use Interview Team Option")
		End If
		If worker.interview_trainer = False or run_interview_team_msg = vbYes Then run_by_interview_team = True 		'the script will run the interview only option
	End If
Next

'Looking for BZ Script writers to allow them to select the option.
For each tester in tester_array                         													'looping through all of the testers
	If user_ID_for_validation = tester.tester_id_number and tester.tester_population = "BZ" Then            'If the person who is running the script is a tester
		continue_with_testing_file = MsgBox("The Interview Script has two run options."  & vbCr & vbCr & "Do you want to run the Interview Team - INTERVIEW ONLY - Option?", vbQuestion + vbYesNo, "Use Interview Team Option")
		If continue_with_testing_file = vbYes Then run_by_interview_team = True
	End If
Next

interview_started_time = time
MFIP_orientation_assessed_and_completed = False

msg_what_script_does_btn 		= 101
msg_save_your_work_btn 			= 102
msg_script_interaction_btn 		= 103
msg_show_instructions_btn 		= 104
msg_script_messaging_btn 		= 105
msg_show_quick_start_guide_btn 	= 106
msg_show_faq_btn 				= 107
interpreter_servicves_btn 		= 108
hsr_manual_interview_btn		= 109
sir_snap_interview				= 110
run_interview_summary_btn		= 99
switch_to_summary = False

form_list = "CAF (DHS-5223)"
form_list = form_list+chr(9)+"HUF (DHS-8107)"
form_list = form_list+chr(9)+"SNAP App for Srs (DHS-5223F)"
form_list = form_list+chr(9)+"MNbenefits"
form_list = form_list+chr(9)+"Combined AR for Certain Pops (DHS-3727)"

when_contact_was_made = date & ", " & time

'Showing the case number dialog
Do
	DO
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 371, 315, "Interview Script Case number dialog"
			EditBox 75, 25, 60, 15, MAXIS_case_number
			DropListBox 75, 45, 145, 15, "Select One:"+chr(9)+form_list, CAF_form
			EditBox 75, 65, 145, 15, worker_signature
			DropListBox 10, 270, 350, 45, "Alert at the time you attempt to save each page of the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
			ButtonGroup ButtonPressed
				OkButton 260, 295, 50, 15
				CancelButton 315, 295, 50, 15
				PushButton 235, 20, 125, 15, "HSR Manual - Interpreter Services", interpreter_servicves_btn
				PushButton 235, 35, 125, 15, "HSR Manual - Interview", hsr_manual_interview_btn
				PushButton 235, 50, 125, 15, "SIR - SNAP Phone Interview Guide", sir_snap_interview
				PushButton 235, 65, 60, 15, "Script Overview", msg_what_script_does_btn
				PushButton 295, 65, 65, 15, "Script How to Use", msg_script_interaction_btn
				If run_by_interview_team = False Then PushButton 10, 160, 120, 15, "Interview Summary", run_interview_summary_btn
				PushButton 240, 200, 120, 15, "More about 'SAVE YOUR WORK'", msg_save_your_work_btn
				PushButton 240, 235, 120, 15, "Details on Dialog Correction", msg_script_messaging_btn
				PushButton 10, 295, 50, 15, "Instructions", msg_show_instructions_btn
				PushButton 60, 295, 70, 15, "Quick Start Guide", msg_show_quick_start_guide_btn
				PushButton 130, 295, 30, 15, "FAQ", msg_show_faq_btn
			GroupBox 5, 10, 220, 75, "Case Information"
			Text 20, 30, 50, 10, "Case number:"
			Text 10, 50, 60, 10, "Actual CAF Form:"
			Text 10, 70, 60, 10, "Worker Signature:"
			GroupBox 230, 10, 135, 75, "Policy and Resources"
			GroupBox 5, 90, 360, 90, "Important Points"
			Text 10, 105, 240, 10, "* * * THIS  SCRIPT  SHOULD  BE  RUN  DURING  THE  INTERVIEW * * *"
			Text 25, 115, 315, 10, "Start this script at the beginning of the interview and use it to record the interview as it happens."
			Text 10, 130, 205, 10, "* Capture info from the form AND info from the conversation."
			If run_by_interview_team = False Then Text 10, 150, 315, 10, "If the interview is already over, we have a temporary option to record the interview information:"
			If run_by_interview_team = True Then Text 10, 165, 120, 10, "Interview Team Functionality Started"
			GroupBox 5, 190, 360, 95, "Script Functionality"
			Text 10, 205, 185, 10, "This script SAVES the information you enter as it runs!"
			Text 10, 215, 345, 10, "IF the script errors, fails, is cancelled, the network goes down. YOU CAN GET YOUR WORK BACK!!!"
			Text 10, 240, 215, 10, "Dialog correction messages can be handled in two different ways."
			Text 10, 255, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed = run_interview_summary_btn Then switch_to_summary = True
		If ButtonPressed > 100 Then
			err_msg = "LOOP"

			If ButtonPressed = msg_what_script_does_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20OVERVIEW.docx"
			If ButtonPressed = msg_script_interaction_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20HOW%20TO%20USE.docx"
			If ButtonPressed = interpreter_servicves_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Interpretive_Services.aspx"
            If ButtonPressed = msg_save_your_work_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20SAVE%20YOUR%20WORK.docx"
			If ButtonPressed = msg_script_messaging_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20SCRIPT%20MESSAGING.docx"

			If ButtonPressed = msg_show_instructions_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW.docx"
			If ButtonPressed = msg_show_quick_start_guide_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20QUICK%20START%20GUIDE.docx"
			If ButtonPressed = msg_show_faq_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NOTES/NOTES%20-%20INTERVIEW%20-%20FAQ.docx"

			If ButtonPressed = hsr_manual_interview_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Interview_Resources.aspx"
			If ButtonPressed = sir_snap_interview Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhssir.cty.dhs.state.mn.us/MAXIS/Documents/SNAP%20Telephone%20Interview%20Guide.pdf"
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

If CAF_form = "MNbenefits" Then page_display = 0

Call check_for_MAXIS(False)
If switch_to_summary = True Then Call run_from_GitHub(script_repository & "notes/interview-summary.vbs" )

function test_edit_access(edit_access_allowed, warning_notice)
	Call back_to_SELF
	Call navigate_to_MAXIS_screen("CASE", "NOTE")
	EMReadScreen err_window_1, 75, 8, 3
	EMReadScreen err_window_2, 75, 9, 3
	EMReadScreen err_window_3, 75, 10, 3
	err_window_1 = ucase(trim(err_window_1))
	err_window_2 = ucase(trim(err_window_2))
	err_window_3 = ucase(trim(err_window_3))
	If err_window_1 = "RPCERROR" or err_window_2 = "RPCERROR" or err_window_3 = "RPCERROR" Then PF3
	PF9
	edit_access_allowed = False
	warning_notice = ""
	EMReadScreen note_prompt, 42, 3, 3
	If note_prompt = "Please enter your note on the lines below:" Then
		edit_access_allowed = True
		PF10
		PF3
		PF3
	Else
		EMReadScreen warning_notice, 75, 24, 2
		warning_notice = trim(warning_notice)
	End If
end function

Call test_edit_access(edit_access_allowed, warning_notice)

If edit_access_allowed = False Then
	If run_by_interview_team = True Then
		edit_access_msg = "                *   ---   *   ---   *   ALERT   *   ---   *   ---   *"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "It appears this case is INACTIVE or a CASE/NOTE cannot be entered."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "An email has been sent to the supervisors to update the case."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "Continue the interview as normal, "
		edit_access_msg = edit_access_msg & vbCr & "the script will save all information."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "Once the supervisor has the case in an editable status, "
		edit_access_msg = edit_access_msg & vbCr & "they will email you that the case is ready."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "-- The script saves all information to be rerun if needed. --"

		email_subject = "Case " & MAXIS_case_number & " - is uneditable"
		email_to_field = "Alexander.Yang@hennepin.us; jeremy.lucca@hennepin.us" '; tammy.coenen@hennepin.us; candace.brown@hennepin.us"
		email_cc_field = ""
		email_body = "Case number: " & MAXIS_case_number & " appears INACTIVE." & vbCr &  "Case is being interviewed by the interview team as of " & now & ". Needs review for REIN or PEND." & vbCr & vbCr & "Warning message when attempting to create a new CASE/NOTE: " & warning_notice
		send_email = True
		If windows_user_ID = "CALO001" Then send_email = False
		If developer_mode = True Then send_email = False
		Call create_outlook_email("", email_to_field, email_cc_field, "", email_subject, 1, False, "", "", False, "", email_body, False, "", send_email)
	Else
		edit_access_msg = "* - * - * ALERT * - * - *"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "It appears you cannot edit this case. "
		edit_access_msg = edit_access_msg & vbCr & vbCr & "You should still continue with the Interview script run"
		edit_access_msg = edit_access_msg & vbCr & "BUT at the end of the script run, it will NOT:"
		edit_access_msg = edit_access_msg & vbCr & "- Enter a CASE/NOTE"
		edit_access_msg = edit_access_msg & vbCr & "- Send a SPEC/MEMO"
		edit_access_msg = edit_access_msg & vbCr & "- Create a Worker Interview Form Document"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "All information captured during the script run will be saved for future access."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "ONCE YOU HAVE ACCESS TO THE CASE:"
		edit_access_msg = edit_access_msg & vbCr & "- Rerun the script for the same Case Number."
		edit_access_msg = edit_access_msg & vbCr & "- Information will be loaded into the script."
		edit_access_msg = edit_access_msg & vbCr & "- The NOTE, MEMO, and Document will be created at the end of this second script run."
		edit_access_msg = edit_access_msg & vbCr & "The details will be saved for 5 days."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "CASE/NOTE Warning Message:"
		edit_access_msg = edit_access_msg & vbCr & warning_notice
		edit_access_msg = edit_access_msg & vbCr & vbCr & "Press CANCEL to stop the script run."
	End If
	' MsgBox edit_access_msg
	If run_by_interview_team = True Then no_acces_msg = MsgBox(edit_access_msg, vbSystemModal + vbExclamation, "No Edit Access to Case")
	If run_by_interview_team = False Then no_acces_msg = MsgBox(edit_access_msg, vbOKCancel + vbSystemModal + vbExclamation, "No Edit Access to Case")
	If no_acces_msg = vbCancel Then
		script_run_lowdown = "edit_access_allowed - " & edit_access_allowed & vbCr & "warning_notice - " & warning_notice & vbCr & vbCr & script_run_lowdown
		call script_end_procedure_with_error_report("~PT: Interview script cancelled at beginning of the script run due to no CASE/NOTE Edit Access.")
	End If
End If
PF3
Call back_to_SELF

Do
	Call navigate_to_MAXIS_screen("STAT", "SUMM")
	EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"
EMReadScreen case_pw, 7, 21, 17


If select_err_msg_handling = "Alert at the time you attempt to save each page of the dialog." Then show_err_msg_during_movement = TRUE
If select_err_msg_handling = "Alert only once completing and leaving the final dialog." Then show_err_msg_during_movement = FALSE

show_known_addr = FALSE
vars_filled = FALSE
membs_found = FALSE

Orig_CAF_form = CAF_form

Call back_to_SELF
Call restore_your_work(vars_filled, membs_found)			'looking for a 'restart' run
'Added the membs_found variable because there were some errors when recording the member information
'the memb array details were all blank. We do not know the source of the error in writing the member detail so all we can do at this point is handle for if it occurs.

Call run_from_GitHub(script_repository & "misc/interview-forms-classes.vbs" )
EMWaitReady 0, 0

caf_version_date = "03/25"
mnbenefits_version_date = "11/16"
huf_version_date = "03/22"
sr_snap_version_date = "04/23"
car_version_date = "04/23"

If vars_filled = True Then

	xmlPath = user_myDocs_folder & "interview_questions_" & MAXIS_case_number & ".xml"

	With (CreateObject("Scripting.FileSystemObject"))

		'Creating an object for the stream of text which we'll use frequently
		Dim objTextStream

		If .FileExists(xmlPath) = True then
			xmlDoc.Async = False

			' Load the XML file
			xmlDoc.load(xmlPath)

			set form_detail_file = objFSO.GetFile(xmlPath)			'create file object
			form_info_modified_date = form_detail_file.DateLastModified						'identify the file create date - which includes date and time

			form_version_date = "Unknown"
			saved_form_version_date = "Unknown"
			If DateDiff("d", #6/23/2025#, form_info_modified_date) >=0 Then
				set node = xmlDoc.SelectSingleNode("//FormVersion")
				saved_form_version_date = node.text
			End If

			set node = xmlDoc.SelectSingleNode("//Name")
			saved_CAF_form = node.text

			correct_version = True
			If Orig_CAF_form = "CAF (DHS-5223)" 							Then selected_form_version_date = caf_version_date
			If Orig_CAF_form = "MNbenefits" 								Then selected_form_version_date = mnbenefits_version_date
			If Orig_CAF_form = "HUF (DHS-8107)" 							Then selected_form_version_date = huf_version_date
			If Orig_CAF_form = "SNAP App for Srs (DHS-5223F)" 				Then selected_form_version_date = sr_snap_version_date
			If Orig_CAF_form = "Combined AR for Certain Pops (DHS-3727)" 	Then selected_form_version_date = car_version_date
			If saved_form_version_date <> selected_form_version_date Then correct_version = False
			If Orig_CAF_form <> saved_CAF_form Then correct_version = False

			If correct_version = False Then
				not_restored_msg = MsgBox("The form or version does not match the saved information." & vbCr & vbCr & "INFORMATION HAS NOT BEEN RESTORED."& vbCr & vbCr &_
										  "Form Selected: " & Orig_CAF_form & vbCr & "Version Supported: " & selected_form_version_date & vbCr & vbCr &_
										  "Saved Form: " & saved_CAF_form & vbCr & "Version: " & saved_form_version_date & vbCr & vbCr &_
										  "Press 'Cancel' to stop the script." & vbCr & "(Most helpful if you selected the wrong form.)", vbCritical + vbOkCancel, "Details could not be Restored")
				If not_restored_msg = vbCancel then script_end_procedure_with_error_report("")
			End If

			If correct_version = True Then
				CAF_form = saved_CAF_form
				form_version_date = saved_form_version_date
				set node = xmlDoc.SelectSingleNode("//DHSNumber")
				form_number = node.text

				set node = xmlDoc.SelectSingleNode("//numbOfQuestions")
				numb_of_quest = node.text
				numb_of_quest = numb_of_quest * 1
				set node = xmlDoc.SelectSingleNode("//lastPageOfQuestions")
				last_page_of_questions = node.text

				set question_nodes = xmlDoc.SelectNodes("//question")
				question_num = 0

				for each node in question_nodes
					item_info_length = 0
					ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "Combined AR for Certain Pops (DHS-3727)"
					Set FORM_QUESTION_ARRAY(question_num) = new form_questions

					call FORM_QUESTION_ARRAY(question_num).restore_info(node)
					FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
					FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
					If FORM_QUESTION_ARRAY(question_num).prefil_btn <> "" Then FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num
					If FORM_QUESTION_ARRAY(question_num).detail_array_exists = True Then FORM_QUESTION_ARRAY(question_num).add_to_array_btn	= 3000+question_num
					question_num = question_num + 1
				next
				set xmlDoc = nothing

				ReDim TEMP_INFO_ARRAY(q_last_const, numb_of_quest)
				If CAF_form = "CAF (DHS-5223)" or CAF_form = "MNbenefits" or CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_HOUSING_ARRAY(5)
				' If CAF_form = "CAF (DHS-5223)" or CAF_form = "MNbenefits" or
                If CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_UTILITIES_ARRAY(3)

				For quest = 0 to UBound(FORM_QUESTION_ARRAY)
					TEMP_INFO_ARRAY(form_yn_const, quest) = FORM_QUESTION_ARRAY(quest).caf_answer
					TEMP_INFO_ARRAY(form_write_in_const, quest) = FORM_QUESTION_ARRAY(quest).write_in_info
					TEMP_INFO_ARRAY(intv_notes_const, quest) = FORM_QUESTION_ARRAY(quest).interview_notes
					TEMP_INFO_ARRAY(form_second_yn_const, quest) = FORM_QUESTION_ARRAY(quest).sub_answer
					TEMP_INFO_ARRAY(form_second_ans_const, quest) = FORM_QUESTION_ARRAY(quest).detail_answer
					If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
						If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then
							unea_1_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(0)
							unea_2_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(1)
							unea_3_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(2)
							unea_4_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(3)
							unea_5_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(4)
							unea_6_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(5)
							unea_7_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(6)
							unea_8_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(7)
							unea_9_amt = FORM_QUESTION_ARRAY(quest).item_detail_list(8)
							unea_1_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(0)
							unea_2_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(1)
							unea_3_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(2)
							unea_4_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(3)
							unea_5_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(4)
							unea_6_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(5)
							unea_7_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(6)
							unea_8_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(7)
							unea_9_yn = FORM_QUESTION_ARRAY(quest).item_ans_list(8)
						End If
						If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then
							For i = 0 to UBound(TEMP_HOUSING_ARRAY)
								TEMP_HOUSING_ARRAY(i) = FORM_QUESTION_ARRAY(quest).item_ans_list(i)
							Next
						End If
						If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then
							For i = 0 to UBound(TEMP_UTILITIES_ARRAY)
								TEMP_UTILITIES_ARRAY(i) = FORM_QUESTION_ARRAY(quest).item_ans_list(i)
							Next
						End If
						If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then
							For i = 0 to UBound(TEMP_ASSETS_ARRAY)
								TEMP_ASSETS_ARRAY(i) = FORM_QUESTION_ARRAY(quest).item_ans_list(i)
							Next
						End If
						If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then
							For i = 0 to UBound(TEMP_MSA_ARRAY)
								TEMP_MSA_ARRAY(i) = FORM_QUESTION_ARRAY(quest).item_ans_list(i)
							Next
						End If
						If FORM_QUESTION_ARRAY(quest).info_type = "stwk" Then
							For i = 0 to UBound(TEMP_STWK_ARRAY)
								TEMP_STWK_ARRAY(i) = FORM_QUESTION_ARRAY(quest).item_ans_list(i)
							Next
						End If
					End If
				Next
			End If
		End If
	End With
End If

expedited_screening_on_form = True
If CAF_form = "CAF (DHS-5223)" Then
	CAF_form_name = "Combined Application Form"
	form_number = "5223"
End If
If CAF_form = "HUF (DHS-8107)" Then
	CAF_form_name = "Household Update Form"
	form_number = "8107"
	expedited_screening_on_form = False
End If
If CAF_form = "SNAP App for Srs (DHS-5223F)" Then
	CAF_form_name = "SNAP Application for Seniors"
	form_number = "5223F"
End If
If CAF_form = "MNbenefits" Then
	CAF_form_name = "MNbenefits Web Form"
	form_number = ""
End If
If CAF_form = "Combined AR for Certain Pops (DHS-3727)" Then
	CAF_form_name = "Combined Annual Renewal"
	form_number = "3727"
	expedited_screening_on_form = False
End If

Call navigate_to_MAXIS_screen("MONY", "INQX")
from_mo = right("00" & DatePart("m", DateAdd("yyyy", -4, date)), 2)
from_yr = right(DatePart("yyyy", DateAdd("yyyy", -4, date)), 2)
EMWriteScreen from_mo, 6, 38
EMWriteScreen from_yr, 6, 41
EMWriteScreen CM_plus_1_mo, 6, 53
EMWriteScreen CM_plus_1_yr, 6, 56
EMWriteScreen "X", 9, 50
EMWriteScreen "X", 11, 50
transmit
mony_row = 6
EMReadScreen issued_date, 8, 6, 7
issued_date = trim(issued_date)
If issued_date <> "" Then
	issued_date = DateAdd("d", 0, issued_date)
	EMReadScreen issued_prog, 3, 6, 16
	issued_prog = trim(issued_prog)
End If

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
EMReadScreen worker_id_for_data_table, 7, 21, 14
EMReadScreen case_name_for_data_table, 25, 21, 40
case_name_for_data_table = trim(case_name_for_data_table)
' MsgBox "case_pending - " & case_pending
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
		' MsgBox "CAF_datestamp - " & CAF_datestamp

		If unknown_cash_pending = True Then CASH_on_CAF_checkbox = checked
		If ga_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If msa_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If mfip_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If dwp_status = "PENDING" Then CASH_on_CAF_checkbox = checked
		If grh_status = "PENDING" Then GRH_on_CAF_checkbox = checked
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
	If cash_revw = True Then
		If ga_status = "ACTIVE" Then CASH_on_CAF_checkbox = checked
		If msa_status = "ACTIVE" Then CASH_on_CAF_checkbox = checked
		If mfip_status = "ACTIVE" Then CASH_on_CAF_checkbox = checked
		If dwp_status = "ACTIVE" Then CASH_on_CAF_checkbox = checked
		If grh_status = "ACTIVE" Then GRH_on_CAF_checkbox = checked
	End If
	If snap_revw = True Then SNAP_on_CAF_checkbox = checked

	If unknown_cash_pending = True Then the_process_for_cash ="Application"
	If ga_status = "PENDING" Then the_process_for_cash = "Application"
	If msa_status = "PENDING" Then the_process_for_cash = "Application"
	If mfip_status = "PENDING" Then the_process_for_cash = "Application"
	If dwp_status = "PENDING" Then the_process_for_cash = "Application"
	If snap_status = "PENDING" Then the_process_for_snap = "Application"
	the_process_for_emer = "Application"
End If

Call read_program_history_case_curr(CASH_ever_active, MSA_ever_active, FS_ever_active, MA_ever_active, EMER_ever_active, GRH_ever_active, GA_ever_active, MFIP_ever_active, DWP_ever_active, QMB_ever_active, SLMB_ever_active, CCAP_ever_active, QI1_ever_active, RCA_ever_active, IV_E_ever_active, IMD_ever_active, CASH_currently_active, MSA_currently_active, FS_currently_active, MA_currently_active, EMER_currently_active, GRH_currently_active, GA_currently_active, MFIP_currently_active, DWP_currently_active, QMB_currently_active, SLMB_currently_active, CCAP_currently_active, QI1_currently_active, RCA_currently_active, IV_E_currently_active, IMD_currently_active, CASH_date_closed, MSA_date_closed, FS_date_closed, MA_date_closed, EMER_date_closed, GRH_date_closed, GA_date_closed, MFIP_date_closed, DWP_date_closed, QMB_date_closed, SLMB_date_closed, CCAP_date_closed, QI1_date_closed, RCA_date_closed, IV_E_date_closed, IMD_date_closed, CASH_reason_closed, MSA_reason_closed, FS_reason_closed, MA_reason_closed, EMER_reason_closed, GRH_reason_closed, GA_reason_closed, MFIP_reason_closed, DWP_reason_closed, QMB_reason_closed, SLMB_reason_closed, CCAP_reason_closed, QI1_reason_closed, RCA_reason_closed, IV_E_reason_closed, IMD_reason_closed, active_spans_array)
PF3

snap_closed_in_past_30_days = False
snap_closed_in_past_4_months = False
If FS_ever_active = True and FS_currently_active = False Then
	If DateDiff("m", FS_date_closed, date) = 0 Then snap_closed_in_past_30_days = True
	If DateDiff("d", FS_date_closed, date) < 31 Then snap_closed_in_past_30_days = True
	If DateDiff("m", FS_date_closed, date) =< 4 Then snap_closed_in_past_4_months = True
End If
grh_closed_in_past_30_days = False
grh_closed_in_past_4_months = False
If GRH_ever_active = True and GRH_currently_active = False Then
	If DateDiff("m", GRH_date_closed, date) = 0 Then grh_closed_in_past_30_days = True
	If DateDiff("d", GRH_date_closed, date) < 31 Then grh_closed_in_past_30_days = True
	If DateDiff("m", GRH_date_closed, date) =< 4 Then grh_closed_in_past_4_months = True
End If
cash1_closed_in_past_30_days = False
cash1_closed_in_past_4_months = False
cash1_recently_closed_program = ""
cash1_date_closed = ""
cash1_closed_reason = ""
cash2_closed_in_past_30_days = False
cash2_closed_in_past_4_months = False
cash2_recently_closed_program = ""
cash2_date_closed = ""
cash2_closed_reason = ""
If DWP_ever_active = True and DWP_currently_active = False Then
	If DateDiff("m", DWP_date_closed, date) = 0 Then cash1_closed_in_past_30_days = True
	If DateDiff("d", DWP_date_closed, date) < 31 Then cash1_closed_in_past_30_days = True
	If DateDiff("m", DWP_date_closed, date) =< 4 Then cash1_closed_in_past_4_months = True
	If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
		cash1_recently_closed_program = "DWP"
		cash1_date_closed = DWP_date_closed
		cash1_closed_reason = DWP_reason_closed
	End If
End If
If MSA_ever_active = True and MSA_currently_active = False Then
	If cash1_recently_closed_program = "" Then
		If DateDiff("m", MSA_date_closed, date) = 0 Then cash1_closed_in_past_30_days = True
		If DateDiff("d", MSA_date_closed, date) < 31 Then cash1_closed_in_past_30_days = True
		If DateDiff("m", MSA_date_closed, date) =< 4 Then cash1_closed_in_past_4_months = True
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
			cash1_recently_closed_program = "MSA"
			cash1_date_closed = MSA_date_closed
			cash1_closed_reason = MSA_reason_closed
		End If
	Else
		If DateDiff("m", MSA_date_closed, date) = 0 Then cash2_closed_in_past_30_days = True
		If DateDiff("d", MSA_date_closed, date) < 31 Then cash2_closed_in_past_30_days = True
		If DateDiff("m", MSA_date_closed, date) =< 4 Then cash2_closed_in_past_4_months = True
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
			cash2_recently_closed_program = "MSA"
			cash2_date_closed = MSA_date_closed
			cash2_closed_reason = MSA_reason_closed
		End If
	End If
End If
If GA_ever_active = True and GA_currently_active = False Then
	If cash1_recently_closed_program = "" Then
		If DateDiff("m", GA_date_closed, date) = 0 Then cash1_closed_in_past_30_days = True
		If DateDiff("d", GA_date_closed, date) < 31 Then cash1_closed_in_past_30_days = True
		If DateDiff("m", GA_date_closed, date) =< 4 Then cash1_closed_in_past_4_months = True
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
			cash1_recently_closed_program = "GA"
			cash1_date_closed = GA_date_closed
			cash1_closed_reason = GA_reason_closed
		End If
	Else
		If DateDiff("m", GA_date_closed, date) = 0 Then cash2_closed_in_past_30_days = True
		If DateDiff("d", GA_date_closed, date) < 31 Then cash2_closed_in_past_30_days = True
		If DateDiff("m", GA_date_closed, date) =< 4 Then cash2_closed_in_past_4_months = True
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
			cash2_recently_closed_program = "GA"
			cash2_date_closed = GA_date_closed
			cash2_closed_reason = GA_reason_closed
		End If
	End If
End If
If MFIP_ever_active = True and MFIP_currently_active = False Then
	If cash1_recently_closed_program = "" Then
		If DateDiff("m", MFIP_date_closed, date) = 0 Then cash1_closed_in_past_30_days = True
		If DateDiff("d", MFIP_date_closed, date) < 31 Then cash1_closed_in_past_30_days = True
		If DateDiff("m", MFIP_date_closed, date) =< 4 Then cash1_closed_in_past_4_months = True
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
			cash1_recently_closed_program = "MFIP"
			cash1_date_closed = MFIP_date_closed
			cash1_closed_reason = MFIP_reason_closed
		End If
	Else
		If DateDiff("m", MFIP_date_closed, date) = 0 Then cash2_closed_in_past_30_days = True
		If DateDiff("d", MFIP_date_closed, date) < 31 Then cash2_closed_in_past_30_days = True
		If DateDiff("m", MFIP_date_closed, date) =< 4 Then cash2_closed_in_past_4_months = True
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
			cash2_recently_closed_program = "MFIP"
			cash2_date_closed = MFIP_date_closed
			cash2_closed_reason = MFIP_reason_closed
		End If
	End If
End If
EMER_active_in_past_12_months = False
If EMER_ever_active = True and EMER_currently_active = False Then
	If DateDiff("m", EMER_date_closed, date) < 12 Then EMER_active_in_past_12_months = True
End If

Do
	DO
		dlg_len = 210
		If run_by_interview_team = True Then dlg_len = 285
		orig_dlg_len = dlg_len
		If snap_closed_in_past_30_days = True or snap_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
		If grh_closed_in_past_30_days = True or grh_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
		If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
		If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then dlg_len = dlg_len + 10
		If issued_date <> "" Then dlg_len = dlg_len + 10
		If dlg_len = orig_dlg_len Then dlg_len = dlg_len + 10

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 326, dlg_len, "Programs to Interview For"
			Text 10, 10, 300, 20, "Record details from the form here for " & CAF_form_name & " being used for this interview:"
			Text 15, 30, 125, 10, "Date form was received in the county:"
			EditBox 140, 25, 45, 15, CAF_datestamp
			Text 200, 30, 110, 10, CAF_form_name
			Text 25, 40, 260, 10, "Active Programs: " & list_active_programs
			Text 25, 50, 260, 10, "Pending Programs: " & list_pending_programs
			GroupBox 10, 65, 265, 30, "Check All Programs Marked on the Form"
			CheckBox 15, 80, 30, 10, "CASH", CASH_on_CAF_checkbox
			CheckBox 55, 80, 35, 10, "SNAP", SNAP_on_CAF_checkbox
			CheckBox 95, 80, 60, 10, "EMERGENCY", EMER_on_CAF_checkbox
			CheckBox 160, 80, 105, 10, "HOUSING SUPPORT (GRH)", GRH_on_CAF_checkbox
			y_pos = 100
			If run_by_interview_team = True Then
				Text 15, y_pos, 180, 10, "About the different programs:"
				Text 20, y_pos+10, 245, 10, "- CASH is a monthly cash benefit."
				Text 20, y_pos+20, 245, 10, "- SNAP is a monthly benefit for the purchase of food items only."
				Text 20, y_pos+30, 245, 10, "- EMERGENCY is a one-time payment to resolve an emergency situation."
				Text 25, y_pos+40, 245, 10, "An example of emergency situation is eviction or utility disconnect."
				Text 20, y_pos+50, 265, 10, "- HOUSING SUPPORT is monthly benefit for people working with an organization"
				Text 25, y_pos+60, 125, 10, "or facility for housing supports."
				y_pos = 175
			End If
			Text 15, y_pos, 200, 10, "Confirm with the resident these were the programs selected."
			Text 15, y_pos+10, 245, 10, "Explain to the resident they can verbally request additional programs to be"
			Text 30, y_pos+20, 125, 10, "assessed while their case is pending. "
			Text 15, y_pos+30, 270, 10, "Explain additionally to the resident they can withdraw their requests at any time."
			y_pos = y_pos + 55
			orig_y_pos = y_pos
			If snap_closed_in_past_30_days = True or snap_closed_in_past_4_months = True Then
				Text 20, y_pos, 285, 10, "SNAP recently closed on " & FS_date_closed & " - " & FS_reason_closed
				y_pos = y_pos + 10
			End If
			If cash1_closed_in_past_30_days = True or cash1_closed_in_past_4_months = True Then
				Text 20, y_pos, 285, 10, cash1_recently_closed_program & " recently closed on " & cash1_date_closed & " - " & cash1_closed_reason
				y_pos = y_pos + 10
			End If
			If cash2_closed_in_past_30_days = True or cash2_closed_in_past_4_months = True Then
				Text 20, y_pos, 285, 10, cash2_recently_closed_program & " recently closed on " & cash2_date_closed & " - " & cash2_closed_reason
				y_pos = y_pos + 10
			End If
			If grh_closed_in_past_30_days = True or grh_closed_in_past_4_months = True Then
				Text 20, y_pos, 285, 10, "GRH/HS recently closed on " & GRH_date_closed & " - " & GRH_reason_closed
				y_pos = y_pos + 10
			End If
			If issued_date <> "" Then
				Text 20, y_pos, 285, 10, "EMER last issued on " & issued_date & " (" & issued_prog & ")"
				y_pos = y_pos + 10
			End If
			If y_pos = orig_y_pos Then
				Text 20, y_pos, 285, 10, "NO RECENT PROGRAM HISTORY TO NOTE"
				y_pos = y_pos + 10
			End If
			GroupBox 10, orig_y_pos-10, 305, y_pos-orig_y_pos+10, "PROGRAM HISTORY"
			Text 15, y_pos+5, 85, 10, "Program Request Notes:"
			EditBox 15, y_pos+15, 300, 15, program_request_notes
			Text 15, y_pos+30, 205, 10, "(Do not document verbal program request or withdrawls here.)"
			ButtonGroup ButtonPressed
				PushButton 235, 45, 80, 15, "Add Verbal Requests", program_requests_btn
				OkButton 210, y_pos+35, 50, 15
				CancelButton 265, y_pos+35, 50, 15
		EndDialog

		err_msg = ""
		Dialog Dialog1
		cancel_confirmation

		cash_other_req_detail = trim(cash_other_req_detail)
	    snap_other_req_detail = trim(snap_other_req_detail)
	    emer_other_req_detail = trim(emer_other_req_detail)

		program_requested = False
		If CASH_on_CAF_checkbox = checked Then program_requested = True
		If GRH_on_CAF_checkbox = checked Then program_requested = True
		If SNAP_on_CAF_checkbox = checked Then program_requested = True
		If EMER_on_CAF_checkbox = checked Then program_requested = True

		If cash_verbal_request = "Yes" Then program_requested = True
		If snap_verbal_request = "Yes" Then program_requested = True
		If emer_verbal_request = "Yes" Then program_requested = True
		If grh_verbal_request = "Yes" Then program_requested = True

		If IsDate(CAF_datestamp) = False Then err_msg = err_msg & vbCr & "* Enter the date of application."
		If program_requested = False Then err_msg = err_msg & vbCr & "* Select which program was requested on the form." & vbCr & "** If there were no programs were marked on the form, the 'Verbal Requests' details can be added using the button."

		If ButtonPressed = program_requests_btn Then
			Call verbal_requests
			err_msg = "LOOP"
		End If

		IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	cash_request = False
	snap_request = False
	emer_request = False
	grh_request = False

	If CASH_on_CAF_checkbox = checked OR cash_verbal_request = "Yes" Then cash_request = True
	If SNAP_on_CAF_checkbox = checked OR snap_verbal_request = "Yes" Then snap_request = True
	If EMER_on_CAF_checkbox = checked OR emer_verbal_request = "Yes" Then emer_request = True
	If GRH_on_CAF_checkbox = checked OR grh_verbal_request = "Yes" Then grh_request = True

	run_process_selection = False
	If cash_request = True Then
		If type_of_cash = "?" or type_of_cash = "" Then run_process_selection = True
		If the_process_for_cash = "Select One..." or the_process_for_cash = "" Then run_process_selection = True
	End If
	If snap_request = True Then
		If the_process_for_snap = "Select One..." or the_process_for_snap = "" Then run_process_selection = True
	End If
	If emer_request = True Then
		If type_of_emer = "?" or type_of_emer = "" Then run_process_selection = True
		If the_process_for_emer = "Select One..." or the_process_for_emer = "" Then run_process_selection = True
	End If
	If grh_request = True Then
		If the_process_for_grh = "Select One..." or the_process_for_grh = "" Then run_process_selection = True
	End If

	If run_process_selection = True Then call program_process_selection

	If snap_request = True AND the_process_for_snap = "Application" Then expedited_determination_needed = True
	If snap_status = "PENDING" Then expedited_determination_needed = True
	If type_of_cash = "Adult" Then family_cash_case_yn = "No"
	If type_of_cash = "Family" Then family_cash_case_yn = "Yes"

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

If vars_filled = TRUE Then show_known_addr = TRUE		'This is a setting for the address dialog to see the view

Call convert_date_into_MAXIS_footer_month(CAF_datestamp, MAXIS_footer_month, MAXIS_footer_year)
original_footer_month = MAXIS_footer_month
original_footer_year = MAXIS_footer_year

'If we already know the variables because we used 'restore your work' OR if there is no case number, we don't need to read the information from MAXIS
If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then
	'Needs to determine MyDocs directory before proceeding.
	intvw_msg_file = user_myDocs_folder & "interview message.txt"

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
End If

If all_the_clients = "" Then Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist

If membs_found = False Then
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
    HH_arrived_date = ""
    HH_arrived_place = ""

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
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "01" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "01 Self"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "02" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "02 Spouse"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "03" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "03 Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "04" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "04 Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "05" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "05 Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "06" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "06 Step Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "08" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "08 Step Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "09" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "09 Step Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "10" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "10 Aunt"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "11" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "11 Uncle"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "12" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "12 Niece"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "13" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "13 Nephew"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "14" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "14 Cousin"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "15" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "15 Grandparent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "16" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "16 Grandchild"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "17" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "17 Other Relative"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "18" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "18 Legal Guardian"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "24" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "24 Not Related"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "25" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "25 Live-in Attendant"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "27" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "27 Unknown"

			If HH_MEMB_ARRAY(id_verif, clt_count) = "BC" Then HH_MEMB_ARRAY(id_verif, clt_count) = "BC - Birth Certificate"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "RE" Then HH_MEMB_ARRAY(id_verif, clt_count) = "RE - Religious Record"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DL - Drivers License/ST ID"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DV - Divorce Decree"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AL - Alien Card"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AD" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AD - Arrival//Depart"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DR" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DR - Doctor Stmt"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "PV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "PV - Passport/Visa"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "OT" Then HH_MEMB_ARRAY(id_verif, clt_count) = "OT - Other Document"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "NO" Then HH_MEMB_ARRAY(id_verif, clt_count) = "NO - No Ver Prvd"

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

			If HH_MEMB_ARRAY(interpreter, clt_count) = "Y" Then HH_MEMB_ARRAY(interpreter, clt_count) = "Yes"
			If HH_MEMB_ARRAY(interpreter, clt_count) = "N" Then HH_MEMB_ARRAY(interpreter, clt_count) = "No"
			If HH_MEMB_ARRAY(alias_yn, clt_count) = "Y" Then HH_MEMB_ARRAY(alias_yn, clt_count) = "Yes"
			If HH_MEMB_ARRAY(alias_yn, clt_count) = "N" Then HH_MEMB_ARRAY(alias_yn, clt_count) = "No"
			If HH_MEMB_ARRAY(ethnicity_yn, clt_count) = "Y" Then HH_MEMB_ARRAY(ethnicity_yn, clt_count) = "Yes"
			If HH_MEMB_ARRAY(ethnicity_yn, clt_count) = "N" Then HH_MEMB_ARRAY(ethnicity_yn, clt_count) = "No"

			HH_MEMB_ARRAY(race, clt_count) = trim(HH_MEMB_ARRAY(race, clt_count))
            HH_MEMB_ARRAY(race_a_checkbox, clt_count) = unchecked
            HH_MEMB_ARRAY(race_b_checkbox, clt_count) = unchecked
            HH_MEMB_ARRAY(race_n_checkbox, clt_count) = unchecked
            HH_MEMB_ARRAY(race_p_checkbox, clt_count) = unchecked
            HH_MEMB_ARRAY(race_w_checkbox, clt_count) = unchecked

            If HH_MEMB_ARRAY(race, clt_count) = "Asian" Then                        HH_MEMB_ARRAY(race_a_checkbox, clt_count) = checked
            If HH_MEMB_ARRAY(race, clt_count) = "Black Or African Amer" Then        HH_MEMB_ARRAY(race_b_checkbox, clt_count) = checked
            If HH_MEMB_ARRAY(race, clt_count) = "Amer Indn Or Alaskan Native" Then  HH_MEMB_ARRAY(race_n_checkbox, clt_count) = checked
            If HH_MEMB_ARRAY(race, clt_count) = "Pacific Is Or Native Hawaii" Then  HH_MEMB_ARRAY(race_p_checkbox, clt_count) = checked
            If HH_MEMB_ARRAY(race, clt_count) = "White" Then                        HH_MEMB_ARRAY(race_w_checkbox, clt_count) = checked
            If HH_MEMB_ARRAY(race, clt_count) = "Multiple Races" Then
                PF9
                call write_value_and_transmit("X", 17, 34)
                EMReadScreen race_pop_up_check, 18, 5, 12
                If race_pop_up_check = "X AS MANY AS APPLY" Then
                    EMReadScreen x_a, 1, 7, 12
                    If x_a = "X" Then HH_MEMB_ARRAY(race_a_checkbox, clt_count) = checked
                    EMReadScreen x_b, 1, 8, 12
                    If x_b = "X" Then HH_MEMB_ARRAY(race_b_checkbox, clt_count) = checked
                    EMReadScreen x_n, 1, 10, 12
                    If x_n = "X" Then HH_MEMB_ARRAY(race_n_checkbox, clt_count) = checked
                    EMReadScreen x_p, 1, 12, 12
                    If x_p = "X" Then HH_MEMB_ARRAY(race_p_checkbox, clt_count) = checked
                    EMReadScreen x_w, 1, 14, 12
                    If x_w = "X" Then HH_MEMB_ARRAY(race_w_checkbox, clt_count) = checked
                End If
                EMReadScreen race_pop_up_check, 7, 6, 22
                EMReadScreen memb_mode, 1, 20, 8
                Do while race_pop_up_check <> "* Last:" and memb_mode <> "D"
                    PF10
                    EMReadScreen race_pop_up_check, 7, 6, 22
                    EMReadScreen memb_mode, 1, 20, 8
                Loop
            End If

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

			If HH_MEMB_ARRAY(marital_status, clt_count) = "N" Then HH_MEMB_ARRAY(marital_status, clt_count) = "N  Never Married"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "M" Then HH_MEMB_ARRAY(marital_status, clt_count) = "M  Married Living with Spouse"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "S" Then HH_MEMB_ARRAY(marital_status, clt_count) = "S  Married Living Apart (Sep)"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "L" Then HH_MEMB_ARRAY(marital_status, clt_count) = "L  Legally Seperated"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "D" Then HH_MEMB_ARRAY(marital_status, clt_count) = "D  Divorced"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "W" Then HH_MEMB_ARRAY(marital_status, clt_count) = "W  Widowed"
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

            If clt_count = 0 Then
                HH_arrived_date = HH_MEMB_ARRAY(mn_entry_date, clt_count)
                HH_arrived_place = HH_MEMB_ARRAY(former_state, clt_count)
            Else
                If HH_arrived_date <> HH_MEMB_ARRAY(mn_entry_date, clt_count) or HH_arrived_place <> HH_MEMB_ARRAY(former_state, clt_count)  Then
                    HH_arrived_date = ""
                    HH_arrived_place = ""
                End If
            End If

		End If

		memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)
		If HH_MEMB_ARRAY(fs_pwe, clt_count) = "Yes" Then the_pwe_for_this_case = HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)

		' HH_MEMB_ARRAY(clt_count).intend_to_reside_in_mn = "Yes"

		' ReDim Preserve ALL_ANSWERS_ARRAY(ans_notes, clt_count)
		clt_count = clt_count + 1
	Next
    If HH_arrived_date <> "" Then all_members_listed_notes = "All members arrived in Minnesota on " & HH_arrived_date & " from " & HH_arrived_place & "."

    Call navigate_to_MAXIS_screen("STAT", "TYPE")		'===============================================================================================
    type_row = 6
	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)

        EMReadScreen type_ref_numb, 2, type_row, 3
        Do While type_ref_numb <> HH_MEMB_ARRAY(ref_number, the_members)
            type_row = type_row + 1
            If type_row > 20 Then
                PF8
                EMWaitReady 0, 0
                type_row = 6
            End If
            EMReadScreen type_ref_numb, 2, type_row, 3
        Loop
        EMReadScreen memb_cash, 1, type_row, 28
        ' EMReadScreen memb_hc, 1, type_row, 37
        EMReadScreen memb_snap, 1, type_row, 46
        EMReadScreen memb_emer, 1, type_row, 55
        memb_grh = "N"
        If type_row = 6 Then EMReadScreen memb_grh, 1, type_row, 64

		HH_MEMB_ARRAY(snap_req_checkbox, the_members) = unchecked
        If SNAP_on_CAF_checkbox = checked Then
            If memb_snap = "Y" Then HH_MEMB_ARRAY(snap_req_checkbox, the_members) = checked
		End If
		HH_MEMB_ARRAY(cash_req_checkbox, the_members) = unchecked
        If CASH_on_CAF_checkbox = checked Then
            If memb_cash = "Y" Then HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked
		End If
        If GRH_on_CAF_checkbox = checked Then
            If memb_grh = "Y" Then HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked
		End If
        HH_MEMB_ARRAY(emer_req_checkbox, the_members) = unchecked
        If EMER_on_CAF_checkbox = checked Then
            If memb_emer = "Y" Then HH_MEMB_ARRAY(emer_req_checkbox, the_members) = checked
		End If

        HH_MEMB_ARRAY(none_req_checkbox, the_members) = checked
		If HH_MEMB_ARRAY(snap_req_checkbox, the_members) = checked Then HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked
		If HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked Then HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked
		If HH_MEMB_ARRAY(emer_req_checkbox, the_members) = checked Then HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked

        HH_MEMB_ARRAY(requires_update, the_members) = False
        If HH_MEMB_ARRAY(rel_to_applcnt, the_members) = "01 Self" and (HH_MEMB_ARRAY(id_verif, the_members) = "__" or HH_MEMB_ARRAY(id_verif, the_members) = "NO - No Ver Prvd") Then
            HH_MEMB_ARRAY(requires_update, the_members) = True
        End If

        ssn_info_valid = True
        If trim(HH_MEMB_ARRAY(ssn, the_members)) = "" Then
            ssn_info_valid = False
            If HH_MEMB_ARRAY(ssn_verif, the_members) = "A - SSN Applied For" Then ssn_info_valid = True
            If HH_MEMB_ARRAY(ssn_verif, the_members) = "N - Member Does Not Have SSN" Then ssn_info_valid = True
        End If
        If HH_MEMB_ARRAY(ssn_verif, the_members) = "N - SSN Not Provided" Then ssn_info_valid = False
        If HH_MEMB_ARRAY(none_req_checkbox, the_members) = checked Then ssn_info_valid = True
        If ssn_info_valid = False Then HH_MEMB_ARRAY(requires_update, the_members) = True

        If HH_MEMB_ARRAY(citizen, the_members) = "No" and HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked Then
            If trim(HH_MEMB_ARRAY(imig_status, the_members)) = "" Then
                HH_MEMB_ARRAY(requires_update, the_members) = True
            End If
            If (HH_MEMB_ARRAY(clt_has_sponsor, the_members) = "?" or HH_MEMB_ARRAY(clt_has_sponsor, the_members) = "") and HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked Then
                HH_MEMB_ARRAY(requires_update, the_members) = True
            End If
        End If

		HH_MEMB_ARRAY(clt_has_sponsor, the_members) = ""
		HH_MEMB_ARRAY(client_verification, the_members) = ""
		HH_MEMB_ARRAY(client_verification_details, the_members) = ""
		HH_MEMB_ARRAY(client_notes, the_members) = ""
		HH_MEMB_ARRAY(imig_status, the_members) = ""
	Next
End If

If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then

	'Now we gather the address information that exists in MAXIS
    Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, reservation_name, mail_line_one, mail_line_two, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, address_change_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_one_type, phone_two_type, phone_three_type, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

	resi_line_one = ""
    resi_line_two = ""
    mail_line_one = ""
    mail_line_two = ""

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
	Call check_for_MAXIS(False)

	oExec.Terminate()
End If

'Giving the buttons specific unumerations so they don't think they are eachother
next_btn					= 100
' back_btn					= 1010
update_information_btn		= 5020
save_information_btn		= 5030
clear_mail_addr_btn			= 5040
clear_phone_one_btn			= 5041
clear_phone_two_btn			= 5042
clear_phone_three_btn		= 5043
add_person_btn				= 5050
clear_job_btn				= 1100
open_r_and_r_btn 			= 1200
cover_letter_btn            = 1250
caf_page_one_btn			= 1300
caf_addr_btn				= 1400
caf_membs_btn				= 1500
caf_q_pg_4_btn				= 1600
caf_q_pg_5_btn				= 1650
caf_q_pg_6_btn				= 1700
caf_q_pg_7_btn				= 1750
caf_q_pg_8_btn				= 1800
caf_q_pg_9_btn				= 1850
caf_q_pg_10_btn				= 1900
caf_q_pg_11_btn				= 1950

caf_qual_q_btn				= 2200
caf_last_page_btn			= 2300
finish_interview_btn		= 2400
exp_income_guidance_btn 	= 2500
discrepancy_questions_btn	= 2600
emer_questions_btn 			= 2605
open_hsr_manual_transfer_page_btn = 2610
incomplete_interview_btn	= 2700
verif_button				= 2800
q_12_all_no_btn				= 2900
q_14_all_no_btn				= 3000
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
program_requests_btn 		= 911

msg_mfip_orientation_btn		= 930
cm_05_12_12_06_btn				= 931
cm_28_12_btn					= 932
open_dhs_4163_btn				= 933
open_dhs_3477_btn				= 934
open_dhs_3323_btn				= 935
open_dhs_3366_btn				= 936
open_dhs_bulletin_21_11_01_btn	= 937
open_dhs_1826_btn				= 938
open_hsr_manual_btn				= 939
mfip_orientation_word_doc_btn	= 940
emps_update_complete_btn		= 941

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
clear_verifs_btn				= 831

open_r_and_r_btn				= 700
accounting_service_desk_btn		= 701
accounting_in_hsr_manual_btn	= 702
open_ebt_brochure_btn			= 703
open_npp_doc					= 704
open_IEVS_doc					= 705
open_appeal_rights_doc			= 706
open_civil_rights_rights_doc	= 707
open_program_info_doc			= 708
open_DV_doc						= 709
open_disa_doc					= 710
open_cs_2647_doc				= 711
open_cs_2929_doc				= 712
open_cs_3323_doc				= 713
open_cs_3393_doc				= 714
open_cs_3163B_doc				= 715
open_cs_2338_doc				= 716
open_cs_5561_doc				= 717
open_cs_2961_doc				= 718
open_cs_2887_doc				= 719
open_cs_3238_doc				= 720
open_cs_2625_doc				= 721
explain_six_month_rept			= 722
explain_change_rept				= 723
explain_monthly_rept			= 724
open_cs_2707_doc				= 725
open_cs_7635_doc				= 726

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
err_selected_memb = ""
pick_a_client = replace(all_the_clients, "Select or Type", "Select One...")


pg_4_label = ""
pg_5_label = ""
pg_6_label = ""
pg_7_label = ""
pg_8_label = ""
pg_9_label = ""
pg_10_label = ""
pg_11_label = ""
For quest = 0 to UBound(FORM_QUESTION_ARRAY)
	' MsgBox "dialog page number - " &  FORM_QUESTION_ARRAY(quest).dialog_page_numb & vbCr & "number - " & FORM_QUESTION_ARRAY(quest).number & vbCr & vbCr & "CAF Answer - " & FORM_QUESTION_ARRAY(quest).caf_answer & vbCr & "Write In - " & FORM_QUESTION_ARRAY(quest).write_in_info & vbCr & "Interview Notes - " & FORM_QUESTION_ARRAY(quest).interview_notes
	If pg_4_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 4 Then pg_4_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	If pg_5_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 5 Then
		If right(pg_4_label, 1) = " " Then pg_4_label = pg_4_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_5_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_6_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 6 Then
		If right(pg_5_label, 1) = " " Then pg_5_label = pg_5_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_6_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_7_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 7 Then
		If right(pg_6_label, 1) = " " Then pg_6_label = pg_6_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_7_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_8_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 8 Then
		If right(pg_7_label, 1) = " " Then pg_7_label = pg_7_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_8_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_9_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 9 Then
		If right(pg_8_label, 1) = " " Then pg_8_label = pg_8_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_9_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_10_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 10 Then
		If right(pg_9_label, 1) = " " Then pg_9_label = pg_9_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_10_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_11_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 11 Then
		If right(pg_10_label, 1) = " " Then pg_10_label = pg_10_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_11_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If

Next
If right(pg_4_label, 1) = " " Then
	pg_4_label = pg_4_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 5
End If
If right(pg_5_label, 1) = " " Then
	pg_5_label = pg_5_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 6
End If
If right(pg_6_label, 1) = " " Then
	pg_6_label = pg_6_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 7
End If
If right(pg_7_label, 1) = " " Then
	pg_7_label = pg_7_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 8
End If
If right(pg_8_label, 1) = " " Then
	pg_8_label = pg_8_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 9
End If
If right(pg_9_label, 1) = " " Then
	pg_9_label = pg_9_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 10
End If
If right(pg_10_label, 1) = " " Then
	pg_10_label = pg_10_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 11
End If
If right(pg_11_label, 1) = " " Then
	pg_11_label = pg_11_label & FORM_QUESTION_ARRAY(numb_of_quest).number
	start_at = 12
End If
show_qual = start_at
emergency_questions = start_at + 1
discrepancy_questions = start_at + 2
expedited_determination = start_at + 3
show_pg_last = start_at + 4

Call navigate_to_MAXIS_screen("STAT", "MEMB")
If vars_filled = TRUE and membs_found = False Then MsgBox "The script was able to fill in the case number but not the member information." & vbCr & vbCr & "BE SURE TO REVIEW PERSON DETAILS", vbExclamation, "Member Information Not Found"

interview_questions_clear = False
leave_loop = False
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

                'This sets the button pressed to a non-button incase 'Cancel' is pressed but the script is not actually cancelled.
                'This is important because otherwise the script thinks the button is a non-existant 'prefil_btn' from the forms class
	            If ButtonPressed = 0 then ButtonPressed = 15000

				Call review_information
				Call assess_caf_1_expedited_questions(expedited_screening)
				Call review_for_discrepancies
				Call verification_dialog
				Call check_for_errors(interview_questions_clear)

				If ButtonPressed = interpreter_servicves_btn Then
                    run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://itwebpw026/content/forms/af/_internal/hhs/human_services/initial_contact_access/AF10196.html"
                Else
                    Call display_errors(err_msg, False, show_err_msg_during_movement)
                End If
				If run_by_interview_team = True and page_display = expedited_determination and err_msg = "" Then Call determine_calculations(exp_det_income, exp_det_assets, exp_det_housing, exp_det_utilities, calculated_resources, calculated_expenses, calculated_low_income_asset_test, calculated_resources_less_than_expenses_test, is_elig_XFS)

				If snap_status <> "ACTIVE" Then Call evaluate_for_expedited(intv_app_month_income, intv_app_month_asset, intv_app_month_housing_expense, intv_exp_pay_heat_checkbox, intv_exp_pay_ac_checkbox, intv_exp_pay_electricity_checkbox, intv_exp_pay_phone_checkbox, app_month_utilities_cost, app_month_expenses, case_is_expedited)

			Loop until err_msg = ""

			call dialog_movement

		Loop until leave_loop = TRUE
		If ButtonPressed <> incomplete_interview_btn Then
			proceed_confirm = MsgBox("Have you completed the Interview?" & vbCr & vbCr &_
									"Once you proceed from this point, there is no opportunity to change information that will be entered in CASE/NOTE or in the Interview Notes PDF." & vbCr & vbCr &_
									"Following this point is only check eDRS and Forms Review." & vbCr & vbCr &_
									"Press 'No' now if you have additional notes to make or information to review/enter. This will bring you back to the main dailogs." & vbCr &_
									"Press 'Yes' to confinue to the final part of the interivew (forms)." & vbCr &_
									"Press 'Cancel' to end the script run.", vbYesNoCancel+ vbQuestion, "Confirm Interview Completed")
			If proceed_confirm = vbCancel then cancel_confirmation
            If proceed_confirm = vbNo then leave_loop = False
		End If
		If ButtonPressed = incomplete_interview_btn Then proceed_confirm = vbYes
	Loop Until proceed_confirm = vbYes
	Call check_for_password(are_we_passworded_out)
    leave_loop = False                              'resetting the leave_loop variable in case the dialog is looped through again.
Loop until are_we_passworded_out = FALSE

If run_by_interview_team = True and developer_mode = False Then
	'This is the early tracking XML. It is deleted when the script hits the main tracking file at the end
	Set xmlTracDoc = CreateObject("Microsoft.XMLDOM")
	xmlTracPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_started_" & MAXIS_case_number & ".xml"

	xmlTracDoc.async = False

	Set root = xmlTracDoc.createElement("interview")
	xmlTracDoc.appendChild root

	Set element = xmlTracDoc.createElement("ScriptRunDate")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(date)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ScriptRunTime")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(time)
	element.appendChild info

	Set element = xmlTracDoc.createElement("WorkerName")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(worker_name)
	element.appendChild info

	Set element = xmlTracDoc.createElement("WindowsUserID")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(windows_user_ID)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CaseNumber")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(MAXIS_case_number)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CaseBasket")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(case_pw)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewDate")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(interview_date)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash_request)
	element.appendChild info

	If cash_request = True Then
		Set element = xmlTracDoc.createElement("TypeOfCASH")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(type_of_cash)
		element.appendChild info
	End If

	Set element = xmlTracDoc.createElement("GRHRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(grh_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("SNAPRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(snap_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("EMERRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(emer_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpeditedDetermination")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(is_elig_XFS)
	element.appendChild info

	xmlTracDoc.save(xmlTracPath)

	Set xml = CreateObject("Msxml2.DOMDocument")
	Set xsl = CreateObject("Msxml2.DOMDocument")

	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	txt = Replace(fso.OpenTextFile(xmlTracPath).ReadAll, "><", ">" & vbCrLf & "<")
	stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
	"<xsl:output method=""xml"" indent=""yes""/>" & _
	"<xsl:template match=""/"">" & _
	"<xsl:copy-of select="".""/>" & _
	"</xsl:template>" & _
	"</xsl:stylesheet>"

	xsl.loadXML stylesheet
	xml.loadXML txt

	xml.transformNode xsl

	xml.Save xmlTracPath
End if

phone_droplist = "Select or Type"
If phone_one_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_one_number
If phone_two_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_two_number
If phone_three_number <> "" Then phone_droplist = phone_droplist+chr(9)+phone_three_number
phone_droplist = phone_droplist+chr(9)+phone_number_selection

If signature_detail = "Accepted Verbally" or second_signature_detail = "Accepted Verbally" Then
	If verbal_sig_date = "" or verbal_sig_time = "" or verbal_sig_phone_number = "" Then
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 246, 200, "Verbal Signature Record"
		If signature_detail = "Accepted Verbally" Then Text 20, 20, 185, 10, "MEMB " & signature_person
		If second_signature_detail = "Accepted Verbally" Then Text 20, 30, 185, 10, "MEMB " & second_signature_person
		Text 10, 10, 115, 10, "Verbal Signature Accepted for:"
		Text 20, 50, 190, 20, "To record a verbal signature the date, time and resident phone number needs to be recorded. "
		Text 20, 75, 105, 10, "Signature was accepted at:"
		Text 25, 95, 20, 10, "Date: "
		Text 25, 115, 20, 10, "Time: "
		EditBox 50, 90, 50, 15, verbal_sig_date
		EditBox 50, 110, 50, 15, verbal_sig_time
		Text 20, 140, 85, 10, "Resident Phone Number:"
		ComboBox 110, 135, 95, 45, phone_droplist, verbal_sig_phone_number
		Text 10, 160, 220, 30, "Based on POLI/TEMP 02.05.25 all information here is needed to document the verbal signature. Details will be entered in CASE/NOTE and the WIF in ECF. "
		ButtonGroup ButtonPressed
			OkButton 190, 180, 50, 15
		EndDialog

		Do
			err_msg = ""
			dialog Dialog1
			cancel_confirmation

			If IsDate(verbal_sig_date) = False Then err_msg = err_msg & vbCr & "* Enter the date you accepted the verbal signature."
			If IsDate(verbal_sig_time) = True Then
				verbal_sig_time = FormatDateTime(verbal_sig_time, 3)
				If InStr(verbal_sig_time, ":") = 0 Then err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
				verbal_sig_time = replace(verbal_sig_time, ":00 ", " ")
			Else
				err_msg = err_msg & vbCr & "* The time information does not appear to be a valid time, review and update."
			End If
			If verbal_sig_phone_number = "" or verbal_sig_phone_number = "Select or Type" Then err_msg = err_msg & vbCr & "* Phone number detail is required."

			If err_msg <> "" Then MsgBox "*****     NOTICE     *****" & vbCr & "Please resolve to continue:" & vbCr & err_msg
		Loop until err_msg = ""
		save_your_work
	End If
End If

If run_by_interview_team = True Then
	determined_income = exp_det_income
	determined_assets = exp_det_assets
	determined_shel = exp_det_housing
	determined_utilities = exp_det_utilities
End if

Call check_for_MAXIS(False)

If need_to_update_addr = "True" then
	need_to_update_addr = True
	If addr_verif = "__" OR addr_verif = "Blank" Then addr_verif = "OT"
End If
If need_to_update_addr = "False" then need_to_update_addr = False


If need_to_update_addr = "True" Then
	If IsDate(address_change_date) = False Then address_change_date = CAF_datestamp
	Call access_ADDR_panel("WRITE", notes_on_address, resi_line_one, resi_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, reservation_name, mail_line_one, mail_line_two, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, address_change_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_one_type, phone_two_type, phone_three_type, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
End If
MAXIS_background_check		'Making sure we aren't stuck in background before Hest function is run

'Units can have either the heat/air standard utility deduction OR one or both of the electric and phone standard utility deduction(s) when they are not responsible for heating and/or cooling expenses.
choice_date = CAF_datestamp
hest_update_heat_ac 	= FALSE
hest_update_electric 	= FALSE
hest_update_phone 		= FALSE

If question_15_heat_ac_yn = "Yes" Then 	'If heat/ac is paid for then set phone and electric to blank because we can only claim heat/ac
	hest_update_heat_ac = TRUE
	prosp_heat_ac_yn = "Y"
	prosp_electric_yn = " "
	prosp_electric_units = "  "
	prosp_phone_yn = " "
	prosp_phone_units = "  "

Else 									'If heat/ac is not paid for, then electric and/or phone can be claimed
	prosp_heat_ac_yn = "N"

	If question_15_electricity_yn = "Yes" Then
		hest_update_electric = TRUE
		prosp_electric_yn = "Y"

	Else
		prosp_electric_yn = "N"
	End If

	If question_15_phone_yn = "Yes" Then
		hest_update_phone = TRUE
		prosp_phone_yn = "Y"
	Else
		prosp_phone_yn = "N"
	End If
End If

'If one of the utlitlies is "Yes" Then we will update HEST
If question_15_heat_ac_yn = "Yes" OR question_15_electricity_yn = "Yes" OR question_15_phone_yn = "Yes" Then
	' msgbox "update hest" & vbcr & "persons paying: " & all_persons_paying & vbcr & "heat/ac: " & hest_update_heat_ac & vbcr & "electric: " & hest_update_electric & vbcr & "phone: " & hest_update_phone & vbcr & "retro_heat_ac_yn" & retro_heat_ac_yn & vbcr & "retro_electric_yn" & retro_electric_yn & vbcr & "retro_phone_yn" & retro_phone_yn
	Call access_HEST_panel("WRITE", all_persons_paying, choice_date, actual_initial_exp, retro_heat_ac_yn, retro_heat_ac_units, retro_heat_ac_amt, retro_electric_yn, retro_electric_units, retro_electric_amt, retro_phone_yn, retro_phone_units, retro_phone_amt, prosp_heat_ac_yn, prosp_heat_ac_units, prosp_heat_ac_amt, prosp_electric_yn, prosp_electric_units, prosp_electric_amt, prosp_phone_yn, prosp_phone_units, prosp_phone_amt, total_utility_expense)
End If

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
			If run_by_interview_team = False Then
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
			End If
			If run_by_interview_team = True Then
				BeginDialog Dialog1, 0, 0, 436, 140, "Interview Incomplete"
					EditBox 10, 70, 420, 15, interview_incomplete_reason
					EditBox 10, 100, 420, 15, incomplete_interview_notes
					ComboBox 100, 120, 85, 45, phone_droplist, phone_number_selection
					ButtonGroup ButtonPressed
						OkButton 380, 120, 50, 15
					GroupBox 5, 10, 425, 40, "Incompleting an Interview"
					Text 15, 20, 405, 25, "We make every attempt to complete the entire interview requirement when we are in contact with the resident. Sometimes this becomes impossible and if we are unable to gather all required information, we must INCOMPLETE the interview. Every attempt should be made to complete the interview first."
					Text 10, 60, 120, 10, "Reason the Interview is Incomplete"
					Text 10, 90, 75, 10, "Additional Notes"
					Text 10, 125, 90, 10, "Phone Number (if known):"
				EndDialog
			End If

			dialog Dialog1
			cancel_confirmation

			interview_incomplete_reason = trim(interview_incomplete_reason)
			incomplete_interview_notes = trim(incomplete_interview_notes)

			If interview_incomplete_reason = "" Then err_msg = err_msg & vbCr & "* Explain why the interview is incomplete."

			If err_msg <> "" Then MsgBox "*****     NOTICE     *****" & vbCr & "Please resolve to continue:" & vbCr & err_msg

		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False

	Call check_for_MAXIS(False)

	If create_incomplete_doc_checkbox = checked Then

	End If

	If create_incomplete_note_checkbox = checked Then
		Call write_verification_CASE_NOTE(create_verif_note)
		Call write_interview_CASE_NOTE
		PF3
	End If

	xmlSavePath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_started_" & MAXIS_case_number & ".xml"
	If ObjFSO.FileExists(xmlSavePath) Then objFSO.DeleteFile(xmlSavePath)

	Call start_a_blank_case_note

	If run_by_interview_team = False Then
		Call write_variable_in_CASE_NOTE("INTERVIEW INCOMPLETE - Attempt made but additional details needed")

		Call write_variable_in_CASE_NOTE("Interview attempted on: " & interview_date)
		If create_incomplete_doc_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Document added to Case File with information that was gathered during this partial interview.")
		If create_incomplete_note_checkbox = checked Then Call write_variable_in_CASE_NOTE("* Previous CASE:NOTE has details of information what was gathered during this partial interview.")
		Call write_bullet_and_variable_in_CASE_NOTE("Reason Interview Incomplete", interview_incomplete_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("Additional Notes", incomplete_interview_notes)
	End If
	If run_by_interview_team = True Then
		CALL write_variable_in_CASE_NOTE("Phone Call from " & who_are_we_completing_the_interview_with & " re: INCOMPLETE INTERVIEW")
		If interpreter_information <> "No Interpreter Used" THEN
			CALL write_variable_in_CASE_NOTE("* Contact was made: " & when_contact_was_made & " w/ interpreter: " & interpreter_information)
		Else
			CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
		End if
		If trim(phone_number_selection) <> "Select or Type" then CALL write_bullet_and_variable_in_CASE_NOTE("Phone number", phone_number_selection)
		CALL write_bullet_and_variable_in_CASE_NOTE("Reason for contact", "Interview attempted but could not be completed.")
		CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", "None")
		Call write_bullet_and_variable_in_CASE_NOTE("Reason Interview Incomplete", interview_incomplete_reason)
		Call write_bullet_and_variable_in_CASE_NOTE("Additional Notes", incomplete_interview_notes)

	End If

	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	time_spent = ((timer - start_time) + add_to_time)/60
	time_spent = Round(time_spent, 2)
	end_msg ="INCOMPLETE INTERVIEW run finished." & vbCr & vbCr & "You spent " & time_spent & " minutes on this interview."
	If create_incomplete_doc_checkbox = checked Then end_msg = end_msg & vbCr & " - Doc created to add to ECF."
	If create_incomplete_note_checkbox = checked Then end_msg = end_msg & vbCr & " - NOTE with gathered information created."

	STATS_manualtime = STATS_manualtime + (timer - start_time + add_to_time)
	Call script_end_procedure(end_msg)
End If
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
If edit_access_allowed = True Then
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
End If
save_your_work

If run_by_interview_team = True Then 'R&R Summary for Interview HSRs only
	Do
		Do
			err_msg = ""
			Dialog1 = ""

			BeginDialog Dialog1, 0, 0, 551, 385, "FORMS and INFORMATION Review with Resident"
				Text 10, 10, 460, 10, "Directions: Read the following information to the resident regarding rights and responsibilities."
				Text 25, 30, 500, 25, "You are responsible for reporting changes that might impact benefit eligibility such as employment, income, property, household status, citizenship/immigration status, address, expenses, etc.). For cash/child care assistance, report changes within 10 days of the change. For SNAP, report changes by the 10th of the month following the change. "
				Text 25, 65, 500, 35, "For SNAP 6-month reporting, only report if income exceeds 130% FPG or if Time Limited Recipient and work hours drop below 20 hrs/week. Change reporters must report changes in source of income, change of over $125/month in gross earned income or unearned income, unit composition, residence, housing expense, child support, and if Time Limited Recipient and work hours drop below 20 hrs/week."
				Text 25, 110, 500, 30, "Providing false or incomplete information can lead to loss of benefits/criminal charges. Agencies may verify your information, requiring your consent. Using your benefits acknowledges that you've reported any changes. For child care, you may need to pay a co-payment, additional costs, or provide children's immigration/citizenship documentation; failure to pay or cooperate may end your assistance."
				Text 25, 145, 500, 30, "You have the right to privacy of your information, reapply anytime, receive a copy of your application, know why applications are delayed, know program rules, live where/with whom you choose, and report expenses. For SNAP appeals, you have 90 days to appeal. For Cash/Child Care appeals, appeal within 30 days of receiving notice. Free legal services are available. Discrimination is illegal; if mistreated by a human service agency, file a complaint."

				If snap_case = True Then
					Text 25, 185, 500, 20, "SNAP Applicants Only: SNAP E&T helps you find work or increase earnings, offers opportunities to train for a new career for free and provides support services while working towards your goal. "
					Text 45, 210, 345, 10, "1. Is anyone in the household interested in learning about education, training, or job search assistance?"
					DropListBox 75, 220, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", interested_in_job_assistance_now
					Text 110, 225, 90, 10, "If so list interested names: "
					EditBox 200, 220, 130, 15, interested_names_now
					Text 45, 245, 260, 10, "2. If not, do you think anyone in your household may be interested in the future? "
					DropListBox 75, 255, 30, 15, ""+chr(9)+"No"+chr(9)+"Yes", interested_in_job_assistance_future
					Text 110, 260, 90, 10, "If so list interested names: "
					EditBox 200, 255, 130, 15, interested_names_future
				End If
				ButtonGroup ButtonPressed
					PushButton 465, 365, 80, 15, "Continue", continue_btn
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If snap_case = True Then
				If interested_in_job_assistance_now = "" Then err_msg = err_msg & vbNewLine & "* Complete Question 1: Is anyone in the household interested in learning about education, training, or job search assistance?"
				If interested_in_job_assistance_now = "Yes" AND trim(interested_names_now) = "" Then  err_msg = err_msg & vbNewLine & "* For question 1, enter HH Member names."
				If interested_in_job_assistance_future = "" Then err_msg = err_msg & vbNewLine & "* Complete Question 2: If not, do you think anyone in your household may be interested in the future?"
				If interested_in_job_assistance_future = "Yes" AND trim(interested_names_future) = "" Then  err_msg = err_msg & vbNewLine & "* For question 2, enter HH Member names."
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False
	save_your_work
Else
	'R&R - Start of short and simple R&R for all workers==========================================================

	'DHS4163 DHS3315A
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


	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
				CheckBox 10, 10, 315, 10, "Check here if Cash on an Electronic Benefit (EBT) Card (DHS- 4163) was reviewed", DHS_4163_checkbox
				DropListBox 195, 80, 135, 15, "Select One..."+chr(9)+"Six-Month"+chr(9)+"Change", snap_reporting_type
				EditBox 410, 80, 50, 15, next_revw_month
				ComboBox 195, 95, 135, 45, "Select or Type"+chr(9)+"Yes - I have my card."+chr(9)+"No - I used to but I've lost it."+chr(9)+"No - I never had a card for this case"+chr(9)+case_card_info, case_card_info
				DropListBox 195, 110, 135, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", clt_knows_how_to_use_ebt_card
				ButtonGroup ButtonPressed
					PushButton 290, 170, 145, 15, "HSR Manual - Accounting", accounting_in_hsr_manual_btn
					PushButton 290, 190, 145, 15, "Accounting Service Desk Sharepoint Site", accounting_service_desk_btn
					PushButton 465, 15, 60, 15, "Open DHS4163", open_dhs_4163_btn
					PushButton 465, 365, 80, 15, "Continue", continue_btn
				Text 345, 85, 65, 10, "Your next renewal is "
				Text 25, 40, 135, 10, "If approved, benefit Issuance: "
				Text 25, 115, 160, 10, "Do you know how to use an EBT card?"
				GroupBox 5, 0, 530, 305, ""
				Text 35, 50, 180, 10, "- SNAP " & snap_day_of_issuance & " of the month"
				Text 30, 195, 260, 10, "-Recipients with unstable housing can have the card mailed to a service center"
				Text 35, 60, 185, 10, "- CASH " & cash_day_of_issuance & " of the month"
				Text 25, 155, 100, 10, "First Time Recipients"
				Text 25, 225, 220, 10, "False information may lead to a loss of benefits. Report changes:  "
				Text 25, 270, 440, 10, "Appeal: SNAP within 90 days; Cash/Child care assistance/healthcare within 30 days of notice "
				Text 25, 130, 440, 10, "EBT Card Customer Service (888) 997-2227 or www.ebtEDGE.com"
				Text 25, 85, 155, 10, "This case is subject to which type of reporting?"
				Text 25, 25, 365, 10, "Provided to help families meet their basic needs including: food, shelter, clothing, utilities and transportation."
				Text 30, 235, 440, 10, "-SNAP by 10th next month"
				Text 30, 245, 440, 10, "-Cash assistance/child care assistance within 10 days"
				Text 30, 255, 440, 10, "-Child care provider change requires 15 day prior"
				Text 30, 165, 180, 10, "-By default, receive card by mail"
				Text 25, 100, 160, 10, "Do you already have an EBT card for this case? "
				Text 30, 180, 260, 10, "-In-person pickup must be confirmed with the accounting service desk"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If ButtonPressed = accounting_in_hsr_manual_btn or ButtonPressed = accounting_service_desk_btn or ButtonPressed = open_dhs_4163_btn Then
				err_msg = "LOOP"
				If ButtonPressed = accounting_in_hsr_manual_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Accounting.aspx"
				If ButtonPressed = accounting_service_desk_btn Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-faa/SitePages/Randle-Unit.aspx"
				If ButtonPressed = open_dhs_4163_btn Then run  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
			End If

			If snap_reporting_type = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since you have reviewed SNAP information, select the correct reporting type for this case to ensure the best information is provided to the household."
			If Trim(next_revw_month) = "" Then err_msg = err_msg & vbNewLine & "* Since you have reviewed SNAP information, indicate the next review month for this case."
			If case_card_info = "Select or Type" or trim(case_card_info) = "" Then err_msg = err_msg & vbNewLine & "* Since you have discussed EBT Information, indicate if the resident has an EBT Card for this case."
			If clt_knows_how_to_use_ebt_card = "Select One..." Then err_msg = err_msg & vbNewLine & "* Since you have discussed EBT Information, indicate if the resident knows how to use their EBT Card."

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = False
	save_your_work

	If clt_knows_how_to_use_ebt_card = "No" then
		Do
			Do
				err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"

				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
					Text 10, 5, 160, 10, "REVIEW the information listed here to the resident:"
					CheckBox 10, 20, 270, 10, "Check here if How to Use Your Minnesota EBT Card was reviewed", DHS_3315A_checkbox
					GroupBox 20, 35, 290, 95, "How to get a card:"
					Text 35, 45, 270, 10, "- Your first card is mailed within 2 business days of your benefits being approved"
					Text 35, 55, 130, 10, "- Replacement cards are mailed"
					Text 50, 65, 170, 10, "Call 1-888-997-2227 to request a replacement card"
					Text 50, 75, 170, 10, "Cards take about 5 business days to arrive."
					Text 50, 85, 240, 10, "$2 charge for all replacement cards, which is reduced from your benefit."
					Text 35, 95, 230, 20, "NOTE: If you have cash benefits, you will be issued a card that has your name on it. SNAP only cases do not have names on the EBT card."
					Text 35, 135, 120, 10, "At a store 'point-of-sale' machine."
					Text 35, 145, 75, 10, "At an ATM (Cash Only)"
					Text 35, 155, 140, 10, "At a check cashing business (Cash Only)"
					GroupBox 20, 125, 290, 55, "Where to use your card:"
					Text 35, 185, 135, 10, "- Call customer service at 888-997-2227"
					Text 35, 195, 165, 10, "- Visit your county or tribal human services office"
					Text 35, 205, 195, 10, "- Visit the ebtEDGE cardholder portal www.ebtEDGE.com"
					Text 35, 215, 195, 20, "- Access the ebtEDGE mobile application, www.FISGLOBal.COM/EBTEDGEMOBILE"
					Text 35, 235, 270, 10, "NOTE: 4 failed attepts to enter your PIN locks your card until 12:01 am the next day"
					GroupBox 20, 175, 290, 85, "How to get or change your PIN:"
					GroupBox 20, 255, 290, 100, "Register to receive EBT Information by Text Message"
					Text 35, 265, 135, 10, "1. Go to www.ebtEDGE.com and log in"
					Text 35, 275, 80, 10, "2. Select 'EBT Account'"
					Text 35, 285, 205, 10, "3. Select 'Messaging Registration' under the Account Services menu"
					Text 35, 295, 140, 10, "4. Enter your mobile (cell) phone number."
					Text 35, 305, 230, 10, "5. Check the box next to SMS Balance, then click the 'Update' button."
					Text 35, 315, 190, 10, "6. Use the same mobil number and text for information:"
					Text 45, 325, 135, 10, "- Current Balance (text 'BAL' to 42265)"
					Text 45, 335, 145, 10, "- Last 5 transactions  (text 'MINI' to 42265)"
					GroupBox 310, 35, 210, 320, "General Care/Use"
					Text 325, 50, 80, 10, "Keep your card safe"
					Text 335, 60, 120, 10, "Lost benefits will not be replaced."
					Text 335, 70, 155, 15, "Do not leave your card lying around or lose it, treat it like a debit card or cash."
					Text 325, 100, 110, 10, "Do not throw your card away"
					Text 335, 110, 150, 20, "The same card will be used every month for as long as you have benefits."
					Text 335, 135, 155, 20, "Even if your cases closes and reopens in the future the same card may be used."
					Text 325, 165, 145, 10, "Misuse of your EBT Card is Unlawful"
					Text 330, 175, 160, 20, "- Selling your card or PIN to others may result in criminal charges and your benefits may end."
					Text 330, 195, 165, 20, "- Attempting to buy tobacco products or alcoholic beverages with your EBT Card is considered fraud."
					Text 330, 215, 165, 20, "- Repeated loss of your card may cause a fraud investigation to be opened on you."
					ButtonGroup ButtonPressed
						PushButton 480, 5, 60, 15, "Open DHS3315A", open_ebt_brochure_btn
						PushButton 465, 365, 80, 15, "Continue", continue_btn
				EndDialog

				dialog Dialog1
				cancel_confirmation

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


	'DHS3979 DHS2759 DHS3353 DHS2920
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
				CheckBox 15, 15, 285, 10, "Check here if Notice of Privacy Practices (DHS- 3979) was reviewed", DHS_3979_checkbox
				CheckBox 15, 95, 330, 10, "Check here if Income and Eligibility Verification System (DHS- 2759) was reviewed", DHS_2759_checkbox
				CheckBox 15, 170, 365, 10, "Check here if Appeal Rights and Civil Rights Notice and Complaints (DHS- 3353) was reviewed", DHS_3353_checkbox
				CheckBox 15, 280, 390, 10, "Check here if Program Information for Cash, Food and Child Care Programs (DHS- 2920) was reviewed", DHS_2920_checkbox


				ButtonGroup ButtonPressed
					PushButton 475, 10, 60, 15, "Open DHS3979", open_npp_doc
					PushButton 475, 95, 60, 15, "Open DHS2759", open_IEVS_doc
					PushButton 475, 170, 60, 15, "Open DHS3353", open_appeal_rights_doc
					PushButton 475, 280, 60, 15, "Open DHS2920", open_program_info_doc
					PushButton 465, 365, 80, 15, "Continue", continue_btn
				GroupBox 5, 85, 535, 80, ""
				Text 30, 290, 430, 10, "-If you do not have enough money to meet your basic needs, you can apply for assistance "
				Text 30, 125, 310, 10, "-SSN is required for anyone requesting help. If you don't have one, you must apply for one."
				Text 30, 105, 365, 10, "-Cash assistance, SNAP and MA require income, asset, and health insurance are verified"
				Text 30, 180, 425, 10, "-If you do not agree with a decision made by the agency, you may appeal"
				Text 40, 230, 415, 10, "-Minnesota Department of Human Rights - Metro: (651) 431-3600 or Greater MN: (800) 657-3510"
				Text 30, 115, 365, 10, "-If there are discrepancies, you must respond in writing within 10 days of notification "
				Text 30, 135, 310, 10, "-Child Support Program checks for employment and current benefits "
				GroupBox 5, 160, 535, 115, ""
				Text 30, 190, 455, 10, "-Appeal cash, child care, and health care in writing, within 30 days of notice (good cause extends to within 90 days) "
				Text 30, 200, 455, 10, "-If your benefits stop, you have the right to reapply "
				Text 30, 210, 455, 10, "-If you feel you have been discriminated against by a human service agency you have the right to file a complaint"
				Text 30, 220, 455, 10, "Resources: "
				Text 40, 240, 415, 10, "-US Department of Health and Human Services Office for Civil Rights"
				Text 40, 250, 415, 10, "-US Department of Agriculture"
				GroupBox 5, 270, 535, 80, ""
				Text 30, 300, 430, 10, "-Food and cash programs require an interview"
				Text 30, 310, 430, 10, "-Required proof: who you are, where you live, what family lives with you, your income and assets"
				Text 30, 320, 505, 10, "-How long you've lived in MN, how many people live with you, how much income you/these people receive each month may impact how much you receive"
				Text 30, 330, 505, 10, "-Cash programs include: DWP, MFIP, GA, MSA, GRH, RCA, MN Child Care Assistance Program "
				GroupBox 5, 0, 535, 90, ""
				Text 30, 25, 505, 10, "-Private information helps determine eligibility, you can refuse but it may impact benefits. We use information collected to ensure accurate benefit issuance."
				Text 30, 35, 365, 10, "-SSN required for medical assistance, some financial help, and child support "
				Text 30, 45, 365, 10, "-Information is only shared with agencies/workers on a need to know basis "
				Text 30, 55, 505, 10, "-Private information is only disclosed to those given permission. If you are under 18 information is shared with your parents, unless requested otherwise."
				Text 30, 65, 505, 10, "-Contact MN Department of Human Services in writing if you believe your privacy rights have been violated"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If ButtonPressed = open_npp_doc or ButtonPressed = open_program_info_doc or ButtonPressed = open_appeal_rights_doc or ButtonPressed = open_IEVS_doc  Then
				err_msg = "LOOP"
				If ButtonPressed = open_npp_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
				If ButtonPressed = open_program_info_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-2920-ENG"
				If ButtonPressed = open_appeal_rights_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3353-ENG"
				If ButtonPressed = open_IEVS_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-2759-ENG"
			End If

			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
	save_your_work

	'DHS3477 DHS4133
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
				CheckBox 10, 15, 250, 10, "Check here if Domestic Violence Information (DHS- 3477) was reviewed", DHS_3477_checkbox
				CheckBox 10, 115, 280, 10, "Check here if Do you have a disability? (DHS- 4133) was reviewed", DHS_4133_checkbox

				ButtonGroup ButtonPressed
					PushButton 470, 5, 60, 15, "Open DHS3477", open_dhs_3477_btn
					PushButton 465, 110, 60, 15, "Open DHS4133", open_disa_doc
					PushButton 465, 365, 80, 15, "Continue", continue_btn
				GroupBox 5, 0, 530, 105, ""
				Text 25, 130, 455, 10, "-Assistance is available for accessing services and benefits if you have a disability at no cost"
				Text 25, 25, 490, 10, "-Domestic violence or abuse is what someone says or does over and over again to make you feel afraid or to control you. Services Available:"
				Text 35, 45, 485, 10, "-Minnesota Coalition for Battered Women (866) 289-6177"
				Text 35, 35, 470, 10, "-National Domestic Violence Hotline (800) 799-7233 (TTY: (800) 787-3224"
				Text 35, 55, 480, 10, "-Minnesota Day One Emergency Shelter and Crisis Hotline (800) 233-1111"
				GroupBox 5, 100, 530, 95, ""
				Text 25, 140, 455, 10, "-Disabilities are physical, sensory, or mental impairment such as diabetes, epilepsy, cancer, learning disorders, clinical depression, etc. "
				Text 25, 150, 455, 10, "-Laws like ADA and Human Rights Act protect your rights and ensure you won't be denied benefits due to a disability"
				Text 35, 65, 470, 10, "-Safe At Home (SAH) Program (651) 201-1399 or (866) 723-3035"
				Text 35, 75, 460, 10, "-Vulnerable Adults (800) 333-2433"
			EndDialog

			dialog Dialog1
			cancel_confirmation

			If ButtonPressed = open_DV_doc or ButtonPressed = open_disa_doc Then
				err_msg = "LOOP"
				If ButtonPressed = open_DV_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
				If ButtonPressed = open_disa_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4133-ENG"

			End If
			IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE
	save_your_work

	If family_cash_case_yn = "Yes" Then
		Do
			Do
				err_msg = ""

				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"

					CheckBox 10, 30, 300, 10, "Check here if Reporting Responsibilities for MFIP Households (DHS- 2647) was reviewed", DHS_2647_checkbox
					CheckBox 10, 130, 300, 10, "Check here if Notice of Requirement to Attend MFIP Overview (DHS- 2929) was reviewed", DHS_2929_checkbox
					CheckBox 10, 240, 280, 10, "Check here if Family Violent Referral (DHS- 3323) was reviewed", DHS_3323_checkbox
					ButtonGroup ButtonPressed
						PushButton 470, 20, 60, 15, "Open DHS2647", open_cs_2647_doc
						PushButton 470, 130, 60, 15, "Open DHS2929", open_cs_2929_doc
						PushButton 470, 240, 60, 15, "Open DHS3323", open_cs_3323_doc
					Text 25, 50, 470, 10, "-Changes must be reported on the monthly Household Report form as applicable otherwise on any Report form within 10 days of the change."
					GroupBox 5, 15, 530, 110, ""
					GroupBox 5, 120, 530, 115, ""
					Text 20, 250, 420, 10, "-If you or someone in your home is a victim of domestic abuse the county can help."
					Text 20, 260, 420, 10, "-Resources: National Domestic Violence Hot Line (800) 799-7233 and Legal Aid (888) 354-5522"
					Text 25, 150, 455, 10, "-If you do not attend the meeting, your MFIP grant may be reduced until you do."
					Text 25, 140, 420, 10, "-All MFIP caregivers are required to attend an MFIP overview and participate in Employment Services."
					GroupBox 5, 230, 530, 120, ""
					ButtonGroup ButtonPressed
						PushButton 465, 365, 80, 15, "Continue", continue_btn
					Text 20, 270, 455, 10, "-Victims of domestic abuse on MFIP are exempt from some rules"
					Text 25, 40, 490, 10, "-MFIP cases must report changes to income, assets and household composition"
				EndDialog

				dialog Dialog1
				cancel_confirmation

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
		save_your_work

		If absent_parent_yn = "Yes" Then
			'In cases where there is at least 1 non-custodial parent:
				'Understanding Child Support - A Handbook for Parents (DHS-3393) (PDF).
				'Referral to Support and Collections (DHS-3163B) (PDF). (This is in addition to the Combined Application Form, for EACH non-custodial parent). See 0012.21.03 (Support From Non-Custodial Parents).
				'Cooperation with Child Support Enforcement (DHS-2338) (PDF). See 0012.21.06 (Child Support Good Cause Exemptions).
			'If a non-parental caregiver applies,
				'MFIP Child Only Assistance (DHS-5561) (PDF).
			Do
				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
						CheckBox 20, 10, 335, 10, "Check here if Understanding Child Support - A Handbook for Parents (DHS- 3393) was reviewed ", DHS_3393_checkbox
						CheckBox 15, 95, 270, 10, "Check here if Referral to Support and Collections (DHS- 3163B) was reviewed", DHS_3163B_checkbox
						CheckBox 15, 180, 295, 10, "Check here if Cooperation with Child Support Enforcement (DHS- 2338) was reviewed", DHS_2338_checkbox
						Text 30, 115, 365, 10, "-Understanding Child Support: A Handbook for Parents (DHS- 3393)"
						Text 30, 200, 455, 10, "-If good cause is granted, you do not have to cooperate and your case is closed"
						GroupBox 5, 0, 535, 90, ""
						Text 30, 105, 365, 10, "-Child support agency uses information you provide to collect child support"
						Text 30, 190, 425, 10, "-Cooperating with the child support agency includes providing requested information and attending appointments "
						Text 30, 25, 440, 10, "-Provides support and guidance regarding child support "
						Text 30, 35, 365, 10, "-Every child has the right to financial and emotional support from both parents"
						GroupBox 5, 85, 535, 90, ""
						GroupBox 5, 170, 535, 90, ""



						If relative_caregiver_yn = "Yes" Then
							CheckBox 15, 265, 320, 10, "Check here if Non-Custodial Caregiver - MFIP Child Only Assistance (DHS- 5561) was reviewed", DHS_5561_checkbox
							GroupBox 5, 255, 535, 90, ""
							Text 30, 275, 430, 10, "-Provides information about MFIP for relatives who care for a relative's child"
						End If
						ButtonGroup ButtonPressed
						PushButton 475, 10, 60, 15, "Open DHS3393", open_cs_3393_doc
						PushButton 475, 95, 60, 15, "Open DHS3163B", open_cs_3163B_doc
						PushButton 475, 180, 60, 15, "Open DHS2338", open_cs_2338_doc
						If relative_caregiver_yn = "Yes" Then PushButton 475, 265, 60, 15, "Open DHS5561", open_cs_5561_doc
						PushButton 465, 365, 80, 15, "Continue", continue_btn

					EndDialog

					dialog Dialog1
					cancel_confirmation


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
			save_your_work
		End If
		save_your_work

		If left(minor_caregiver_yn, 3) = "Yes" Then
			'If there is a custodial parent under 20, the
				'Notice of Requirement to Attend School (DHS-2961) (PDF) and
				'Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS-2887) (PDF).
			'If there is a custodial parent under age 18, the
				'MFIP for Minor Caregivers (DHS-3238) (PDF) brochure.
			Do
				Do
					err_msg = ""

					Dialog1 = ""
					BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
						CheckBox 10, 15, 275, 10, "Check here if Notice of Requirement to Attend School (DHS- 2961) was reviewed", DHS_2961_checkbox
						CheckBox 10, 115, 280, 10, "Check here if MFIP Teen Parent Information Brochure (DHS- 2887) was reviewed", DHS_2887_checkbox
						GroupBox 5, 0, 530, 105, ""
						GroupBox 5, 100, 530, 100, ""
						Text 25, 135, 455, 10, "-County human services provides support attendance "
						Text 25, 145, 455, 10, "-Failure to comply without good cause results in 10% or more grant reduction"
						Text 25, 125, 430, 10, "-Teen parents under 20 years old without a high school diploma must attend an approved educational program to qualify for MFIP"
						Text 25, 45, 470, 10, "-Assess needs to support school attendance "
						Text 25, 35, 470, 10, "-Failure to comply without good cause results in loss of MFIP benefits "
						GroupBox 5, 195, 530, 105, ""
						Text 25, 25, 490, 10, "-Required to attend school, unless exempt"


						If minor_caregiver_yn = "Yes - Caregiver is under 18" Then
							CheckBox 10, 210, 280, 10, "Check here if MFIP for Minor Caregivers (DHS- 3238) was reviewed", DHS_3238_checkbox
							Text 20, 225, 505, 20, "You are a minor caregiver if: "
							Text 30, 235, 500, 20, "- You are younger than 18 - You have never been married - You are not emancipated and - You are the parent of a child(ren) living in the same household."
							Text 20, 245, 505, 10, "If you are a minor caregiver, to receive benefits and services, you must be living: "
							Text 30, 255, 465, 20, "- With a parent or with an adult relative caregiver or with a legal guardian or - In an agency-approved living arrangement."
							Text 20, 265, 505, 10, "A social worker must approve any exception(s) to your living arrangement."
						End If

						ButtonGroup ButtonPressed
							PushButton 470, 5, 60, 15, "Open DHS2961", open_cs_2961_doc
							PushButton 465, 110, 60, 15, "Open DHS2887", open_cs_2887_doc
							If minor_caregiver_yn = "Yes - Caregiver is under 18" Then PushButton 465, 205, 60, 15, "Open DHS3238", open_cs_3238_doc
							PushButton 465, 365, 80, 15, "Continue", continue_btn

					EndDialog

					dialog Dialog1
					cancel_confirmation

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
			save_your_work
		End If
	End If
	save_your_work

	If snap_case = True OR pend_snap_on_case = "Yes" OR mfip_status <> "INACTIVE" Then
		'SNAP CASES'
			'Supplemental Nutrition Assistance Program reporting responsibilities (DHS-2625).
			'Facts on Voluntarily Quitting Your Job If You Are on the Supplemental Nutrition Assistance Program (SNAP) (DHS-2707).
			'Work Registration Notice (DHS-7635).
		Do
			Do
				err_msg = ""
				Dialog1 = ""
				BeginDialog Dialog1, 0, 0, 550, 385, "FORMS and INFORMATION Review with Resident"
					CheckBox 10, 10, 385, 10, "Check here if Supplemental Nutrition Assistance Program Reporting Responsibilities (DHS- 2625) was reviewed", DHS_2625_checkbox
					Text 25, 25, 365, 10, snap_reporting_type & " Reporting"

					If snap_reporting_type = "Six-Month" Then
						Text 35, 40, 435, 10, "As a 6 month reporter, you are certified for six months at a time, which means you will have a  review within six months "
						Text 35, 55, 160, 10, "Changes required to report"
						Text 45, 65, 440, 10, "-Income received in any month exceeds 130% FPG for the Household Size"
						Text 45, 75, 440, 10, "-For any ABAWD, a change in work or job activities that cause their hours to fall below 20 hours per week, averaged 80 hours monthly"
						Text 35, 95, 470, 10, "It can be beneficial to report other changes, and we encourage you to do this. Examples include: "
						Text 45, 105, 430, 10, "-Address Change: we communicate via mail and missing mail can cause your benefits to close for lack of response"
						Text 45, 115, 450, 10, "-Decreases in Income: Income is used to determine your benefit amount and any reduction MAY cause your benefit amount to increase"
						Text 45, 125, 450, 10, "-Other Expenses: Child Care, Child Support, sometimes Medical Expenses: Can impact your benefit amount "
						Text 25, 155, 135, 10, "Your next renewal is " & next_revw_month
						Text 25, 170, 245, 10, "Complete the required form and process the month before renewal. "
						Text 25, 185, 440, 10, "Report changes by the 10th of the month following the month of the change"
						Text 25, 200, 445, 20, "SNAP General Work Rules require some household members to accept any job offers and to maintain their current job/hours. If not met, benefits could be decreased/ended. "
						ButtonGroup ButtonPressed
							PushButton 25, 225, 210, 15, "Press here for a list of exemptions from work rules.", exemptions_button
							PushButton 470, 10, 60, 15, "Open DHS2625", open_cs_2625_doc
							PushButton 465, 365, 80, 15, "Continue", continue_btn
						GroupBox 5, 0, 535, 280, ""
					End If
					If snap_reporting_type = "Change" Then
						CheckBox 10, 10, 385, 10, "Check here if Supplemental Nutrition Assistance Program Reporting Responsibilities (DHS- 2625) was reviewed", DHS_2625_checkbox
						ButtonGroup ButtonPressed
							PushButton 25, 235, 210, 15, "Press here for a list of exemptions from work rules.", exemptions_button
							PushButton 470, 10, 60, 15, "Open DHS2625", open_cs_2625_doc
							PushButton 465, 365, 80, 15, "Continue", continue_btn
						Text 25, 165, 135, 10, "Your next renewal is " & next_revw_month
						Text 35, 55, 160, 10, "Change required to report:"
						Text 35, 40, 435, 10, "As a Change Reporter, you typically have a certification period of a year but it could be two years."
						Text 45, 65, 480, 10, "-Change in the source of income, including starting or stopping a job, if the change in employment is accompanied by a change in income."
						Text 25, 180, 245, 10, "Complete the required form and process the month before renewal. "
						Text 25, 25, 365, 10, snap_reporting_type & " Reporting"
						Text 25, 195, 440, 10, "Report changes by the 10th of the month following the month of the change"
						Text 25, 210, 445, 20, "SNAP General Work Rules require some household members to accept any job offers and to maintain their current job/hours. If not met, benefits could be decreased/ended. "
						Text 45, 75, 480, 10, "-A change in more than $125 per month in gross earned income"
						Text 45, 85, 480, 10, "-A change of more than $125 in the amount of unearned income EXCEPT changes related to public assistance "
						Text 45, 95, 480, 10, "-A change in unit composition "
						Text 45, 105, 480, 10, "-A change in residence"
						Text 45, 115, 480, 10, "-A change in housing expense due to residency change "
						Text 45, 125, 480, 10, "-A change in legal obligation to pay child support"
						Text 45, 135, 480, 10, "-For any ABAWD, a change in work or job activities that cause their hours to fall below 20 hours per week, averaged 80 hours monthly. "
						GroupBox 5, 0, 540, 270, ""
					End If
					'If snap_reporting_type = "Monthly" Then
					'	CheckBox 10, 10, 385, 10, "Check here if Supplemental Nutrition Assistance Program Reporting Responsibilities (DHS- 2625) was reviewed", DHS_2625_checkbox
					'	ButtonGroup ButtonPressed
					'		PushButton 20, 185, 210, 15, "Press here for a list of exemptions from work rules.", exemptions_button
					'		PushButton 470, 10, 60, 15, "Open DHS2625", open_cs_2625_doc
					'		PushButton 465, 365, 80, 15, "Continue", continue_btn
					'	Text 25, 115, 135, 10, "Your next renewal is " & next_revw_month
					'	GroupBox 5, 0, 530, 265, ""
					'	Text 40, 40, 435, 20, "As a Monthly Reporter, you are certified for twelve months at a time, which means you will have a review within 12 months. However, the system will close your benefits if the monthly Household Report Form is not received, processed, and all verifications attached. "
					'	Text 25, 130, 245, 10, "Complete the required form and process the month before renewal. "
					'	Text 25, 145, 440, 10, "Report changes by the 10th of the month following the month of the change"
					'	Text 40, 65, 470, 10, "Monthly reporters are required to submit a Household Report Form every month with income and change verifications attached"
					'	Text 25, 160, 445, 20, "SNAP General Work Rules require some household members to accept any job offers and to maintain their current job/hours. If not met, benefits could be decreased/ended. "
					'	Text 40, 80, 470, 20, "The Household Report Form must be answered in its entirety. Any unanswered questions will make the form incomplete and ongoing benefits will not be able to be processed. The form includes all changes that must be reported."
					'End If
				EndDialog


				dialog Dialog1
				cancel_confirmation

				If ButtonPressed = open_cs_2625_doc or ButtonPressed = exemptions_button Then
					err_msg = "LOOP"
					If ButtonPressed = open_cs_2625_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2625-ENG"
					If ButtonPressed = exemptions_button Then call display_exemptions()
				End If

				IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE
		save_your_work
	End If
	save_your_work
	'END HERE WITH R&R CODE ======================================
End If

'Employment Services Registration.

'REPORTING

'Additional Important Information.

'Penalty Warnings.


' Call provide_resources_information(case_number_known, create_case_note, note_detail_array, allow_cancel)
Call provide_resources_information(True, False, note_detail_array, False)
If IsArray(note_detail_array) = True Then
    If IsArray(note_detail_array) = True Then
		all_items_are_blank = True
    	For each note_line in note_detail_array
    		IF note_line <> "" Then	all_items_are_blank = False
		Next
	End If
	If all_items_are_blank = True Then STATS_manualtime = STATS_manualtime + 150
Else
	STATS_manualtime = STATS_manualtime + 150
End If

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

        BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"
            Text 150, 10, 395, 10, "We have finished gathering all the information for the interview. Finish by reviewing this information with the resident."
            GroupBox 10, 20, 530, 325, "CASE INTERVIEW WRAP UP"
            y_pos = 45
            If run_by_interview_team = True Then
                Text 15, 35, 505, 10, "Explain Verifications:"
                Text 20, 45, 505, 10, "If verifications are needed, a request will be sent in the mail. Provide proofs quickly, as they are due in 10 days."
                Text 20, 55, 505, 10, "We can help you obtain these verifications if you have any difficulties. Contact us by phone or come to a service center if you need help."
                Text 15, 75, 460, 10, "Your case will be processed by another worker, there is a possibility they will need to contact you with additional clarifications."
                Text 25, 90, 150, 10, "Confirm the best Phone Number to reach you:"
                ComboBox 175, 85, 85, 45, phone_droplist, phone_number_selection
                Text 270, 90, 170, 10, "Can we leave a detailed message at this number?"
                DropListBox 440, 85, 60, 45, "?"+chr(9)+"Yes"+chr(9)+"No", leave_a_message
                Text 25, 105, 400, 10, "Do you have any questions or requests I can pass on to the processing worker or a program specialist?"
                EditBox 25, 115, 475, 15, resident_questions
                y_pos = 140
            Else
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
            End If
            Text 15, y_pos, 505, 10, "Your address and phone number are our best way to contact you, let us know if these change so you do not miss any notices or requests."
            y_pos = y_pos + 10
            Text 20, y_pos, 505, 10, "Our mail does not forward and missing notices can cause your benefits to end."
            y_pos = y_pos + 20
            Text 15, y_pos, 505, 10, "If you are unsure of program rules and requirements, the forms we reviewed earlier can always be resent, or you can call us with questions."
            y_pos = y_pos + 20
            GroupBox 15, y_pos, 505, 95, "Contact Hennepin County by phone, in person, or online. Ask the resident if they need any more details:"
            y_pos = y_pos + 10
            Text 20, y_pos, 40, 10, "By Phone:"
            Text 60, y_pos, 450, 10, "612-596-1300. The phone lines are open Monday - Friday from 9:00 - 4:00"
            y_pos = y_pos + 10
            Text 20, y_pos, 40, 10, "In person: "
            Text 60, y_pos, 170, 10, "Northwest Human Service Center"
            Text 230, y_pos, 200, 10, "7051 Brooklyn Blvd Brooklyn Center 55429"
            y_pos = y_pos + 10
            Text 60, y_pos, 170, 10, "North Minneapolis Service Center"
            Text 230, y_pos, 200, 10, "1001 Plymouth Ave N Minneapolis 55411"
            y_pos = y_pos + 10
            Text 60, y_pos, 170, 10, "South Minneapolis Human Service Center"
            Text 230, y_pos, 200, 10, "2215 East Lake Street Minneapolis 55407"
            y_pos = y_pos + 10
            Text 60, y_pos, 170, 10, "Health Services Building (Downtown Minneapolis)"
            Text 230, y_pos, 200, 10, "525  Portland Ave S (5th floor) Minneapolis 55415"
            y_pos = y_pos + 10
            Text 60, y_pos, 170, 10, "South Suburban Human Service Center"
            Text 230, y_pos, 200, 10, "9600 Aldrich Ave S Bloomington 55420"
            y_pos = y_pos + 10
            Text 20, y_pos, 40, 10, "Online:"
            Text 60, y_pos, 400, 10, "MNBenefits  at  https://mnbenefits.mn.gov/  -  Use for submitting applications and documents."
            y_pos = y_pos + 10
            Text 60, y_pos, 465, 10, "InfoKeep  at  https://infokeep.hennepin.us/  -  Create a unique sign in to submit documents directly to your case, has a chat functionality."
            y_pos = y_pos + 25
            If run_by_interview_team = True Then Text 15, y_pos, 270, 10, "Summarize any additional case details to pass on to the processing worker:"
            If run_by_interview_team = False Then Text 15, y_pos, 270, 10, "Summarize what is happening with this case:"
            y_pos = y_pos + 10
            EditBox 15, y_pos, 520, 15, case_summary
            Text 10, 370, 220, 10, "Confirm you have reviewed Hennepin County Information Information:"
            DropListBox 230, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Recap Discussed"+chr(9)+"No, I could not complete this", confirm_recap_read
            ButtonGroup ButtonPressed
                PushButton 465, 365, 80, 15, "Interview Completed", continue_btn
        EndDialog

        dialog Dialog1
        cancel_confirmation

        If confirm_recap_read = "Enter confirmation" Then err_msg = err_msg & vbNewLine & "* Indicate if this required information was reviewed with the resident completing the interview."

        If run_by_interview_team = True Then
            resident_questions = trim(resident_questions)
            phone_number_selection = trim(phone_number_selection)
            If phone_number_selection = "" or phone_number_selection = "Select or Type" Then err_msg = err_msg & vbNewLine & "* Enter a phone number to reach the resident at in the case of follow up questions."
            If leave_a_message = "?" Then err_msg = err_msg & vbNewLine & "* Indicate if a detailed message can be left at the phone number provided."
        End If

        If ButtonPressed = verif_button Then
            Call verification_dialog
            err_msg = "LOOP"
        End If

        IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE
save_your_work
Call check_for_MAXIS(False)

edit_access_allowed = False
warning_notice = ""
Call test_edit_access(edit_access_allowed, warning_notice)

If edit_access_allowed = False Then
	If run_by_interview_team = True Then
		edit_access_msg = "                *   ---   *   ---   *   ALERT   *   ---   *   ---   *"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "Inactive case handling is in effect on this case."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "It appears this case is INACTIVE or a CASE/NOTE cannot be entered. The script has NOT:"
		edit_access_msg = edit_access_msg & vbCr & "  - Entered a CASE/NOTE"
		edit_access_msg = edit_access_msg & vbCr & "  - Sent a SPEC/MEMO"
		edit_access_msg = edit_access_msg & vbCr & "  - Created a Worker Interview Form Document"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "All information captured during the script run is saved for future access."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "ONCE THE SUPERVISOR LETS YOU KNOW THE CASE IS READY:"
		edit_access_msg = edit_access_msg & vbCr & "  - Rerun the script for the same Case Number."
		edit_access_msg = edit_access_msg & vbCr & "  - Information will be loaded into the script."
		edit_access_msg = edit_access_msg & vbCr & "  - The NOTE, MEMO, and Document will be created at the end of this second script run."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "Information of the script run is saved for 5 days, check in with your supervisor with any questions."
	Else
		edit_access_msg = "* - * - * SCRIPT RUN ENDED * - * - *"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "It appears you cannot edit this case. "
		edit_access_msg = edit_access_msg & vbCr & "The script has NOT:"
		edit_access_msg = edit_access_msg & vbCr & "- Entered a CASE/NOTE"
		edit_access_msg = edit_access_msg & vbCr & "- Sent a SPEC/MEMO"
		edit_access_msg = edit_access_msg & vbCr & "- Created a Worker Interview Form Document"
		edit_access_msg = edit_access_msg & vbCr & vbCr & "All information captured during the script run is saved for future access."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "ONCE YOU HAVE ACCESS TO THE CASE:"
		edit_access_msg = edit_access_msg & vbCr & "- Rerun the script for the same Case Number."
		edit_access_msg = edit_access_msg & vbCr & "- Information will be loaded into the script."
		edit_access_msg = edit_access_msg & vbCr & "- The NOTE, MEMO, and Document will be created at the end of this second script run."
		edit_access_msg = edit_access_msg & vbCr & "The details will be saved for 5 days."
		edit_access_msg = edit_access_msg & vbCr & vbCr & "CASE/NOTE Warning Message:"
		edit_access_msg = edit_access_msg & vbCr & warning_notice
		' MsgBox edit_access_msg
	End If
	script_run_lowdown = "edit_access_allowed - " & edit_access_allowed & vbCr & "warning_notice - " & warning_notice & vbCr & vbCr & script_run_lowdown
	call script_end_procedure_with_error_report(edit_access_msg)
End If
Call back_to_SELF

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

		'Ensuring the interview date isn't entered on a line/for a program that doesn't have an APPL date. This helps to support the verbal request funcitonality
		EMReadScreen prog_warning, 78, 24, 2
		prog_warning = Trim(UCase(prog_warning))
		Do while prog_warning <> ""
			If InStr(prog_warning, "CASH I ENTRY INVALID") Then
				EMWriteScreen "  ", 6, 55               'CASH I Row
				EMWriteScreen "  ", 6, 58
				EMWriteScreen "  ", 6, 61
			End If
			If InStr(prog_warning, "CASH I ENTRY INVALID") Then
				EMWriteScreen "  ", 7, 55               'CASH II Row
				EMWriteScreen "  ", 7, 58
				EMWriteScreen "  ", 7, 61
			End If
			If InStr(prog_warning, "EMER PROGRAM NOT SELECTED") Then
				EMWriteScreen "  ", 8, 55               'EMER Row
				EMWriteScreen "  ", 8, 58
				EMWriteScreen "  ", 8, 61
			End If
			If InStr(prog_warning, "GRH PROGRAM NOT SELECTED") Then
				EMWriteScreen "  ", 9, 55               'GRH Row
				EMWriteScreen "  ", 9, 58
				EMWriteScreen "  ", 9, 61
			End If
			If InStr(prog_warning, "FS PROGRAM NOT SELECTED") Then
				EMWriteScreen "  ", 10, 55               'SNAP Row
				EMWriteScreen "  ", 10, 58
				EMWriteScreen "  ", 10, 61
			End If

			If prog_warning = "ENTER A VALID COMMAND OR PF-KEY" Then Exit Do

			transmit
			EMReadScreen prog_warning, 78, 24, 2
			prog_warning = Trim(UCase(prog_warning))
		Loop

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
With (CreateObject("Scripting.FileSystemObject"))
	If .FileExists(intvw_done_msg_file) = True then .DeleteFile(intvw_done_msg_file)

	If .FileExists(intvw_done_msg_file) = False then
		Set objTextStream = .OpenTextFile(intvw_done_msg_file, 2, true)

		'Write the contents of the text file
		objTextStream.WriteLine "This interview has been COMPLETED!"
		objTextStream.WriteLine ""
		objTextStream.WriteLine "The interview took " & interview_time & " minutes."
		objTextStream.WriteLine "The script is currently creating your PDF, SPEC/MEMO, and CASE/NOTEs. DO NOT TRY TO TAKE ANY ACTION ON THE COMPUTER WHILE THIS FINISHES."
		objTextStream.WriteLine "Agency Signature is not required on the " & CAF_form & "."
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

Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMReadScreen memb_check, 4, 2, 48
Do While memb_check <> "MEMB"
    Call back_to_SELF
    Call MAXIS_background_check
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
    EMReadScreen memb_check, 4, 2, 48
Loop

Dim CHANGES_ARRAY()
ReDim CHANGES_ARRAY(last_const, 0)	'Defining the changes array to

For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
    ReDim Preserve CHANGES_ARRAY(last_const, the_memb)
    If HH_MEMB_ARRAY(pers_in_maxis, the_memb) = True and HH_MEMB_ARRAY(ignore_person, the_memb) = False Then

        Call navigate_to_MAXIS_screen("STAT", "MEMB")
        EMWriteScreen HH_MEMB_ARRAY(ref_number, the_memb), 20, 76
        transmit

        EMReadscreen curr_last_name, 25, 6, 30
        EMReadscreen curr_first_name, 12, 6, 63
        EMReadscreen curr_mid_initial, 1, 6, 79
        EMReadScreen curr_age, 3, 8, 76

        curr_last_name = trim(replace(curr_last_name, "_", ""))
        curr_first_name = trim(replace(curr_first_name, "_", ""))
        curr_mid_initial = trim(replace(curr_mid_initial, "_", ""))
        curr_age = trim(curr_age)
        If curr_age = "" Then curr_age = 0
        curr_age = curr_age * 1

        If curr_last_name <> HH_MEMB_ARRAY(last_name_const, the_memb)   Then CHANGES_ARRAY(last_name_const, the_memb) = curr_last_name
        If curr_first_name <> HH_MEMB_ARRAY(first_name_const, the_memb) Then CHANGES_ARRAY(first_name_const, the_memb) = curr_first_name
        If curr_mid_initial <> HH_MEMB_ARRAY(mid_initial, the_memb)     Then CHANGES_ARRAY(mid_initial, the_memb) = curr_mid_initial
        If curr_age <> HH_MEMB_ARRAY(age, the_memb)                     Then CHANGES_ARRAY(age, the_memb) = curr_age

        EMReadScreen curr_date_of_birth, 10, 8, 42
        EMReadScreen curr_ssn, 11, 7, 42
        EMReadScreen curr_ssn_verif, 1, 7, 68
        EMReadScreen curr_birthdate_verif, 2, 8, 68
        EMReadScreen curr_gender, 1, 9, 42
        EMReadScreen curr_race, 30, 17, 42
        EMReadScreen curr_spoken_lang, 2, 12, 42
        EMReadScreen curr_written_lang, 2, 13, 42
        EMReadScreen curr_interpreter, 1, 14, 68
        EMReadScreen curr_alias_yn, 1, 15, 42
        EMReadScreen curr_ethnicity_yn, 1, 16, 68

        curr_date_of_birth = replace(curr_date_of_birth, " ", "/")
        curr_ssn = replace(curr_ssn, " ", "-")
        if curr_ssn = "___-__-____" Then curr_ssn = ""
        curr_race = trim(curr_race)

        If curr_date_of_birth <> HH_MEMB_ARRAY(date_of_birth, the_memb)         Then CHANGES_ARRAY(date_of_birth, the_memb) = curr_date_of_birth
        If curr_ssn <> HH_MEMB_ARRAY(ssn, the_memb)                             Then
            ssn_update_attempt = True

            CHANGES_ARRAY(ssn, the_memb) = curr_ssn
            PF9
            numb_only_ssn = replace(replace(curr_ssn, "-", ""), " ", "")
            EMWriteScreen left(numb_only_ssn, 3), 7, 42
            EMWriteScreen mid(numb_only_ssn, 4, 2), 7, 46
            EMWriteScreen right(numb_only_ssn, 4), 7, 49
            EMWriteScreen "P", 7, 68
            curr_ssn_verif = "P"
            transmit

            EMReadScreen memb_check, 4, 2, 48
            If memb_check = "MEMB" Then
                PF10
                EMWaitReady 0, 0
            End If

            attempt_count = 0
            EMReadScreen match_check, 4, 2, 51
            Do While match_check = "MTCH"
                PF3
                EMWaitReady 0, 0
                EMReadScreen match_check, 4, 2, 51
                attempt_count = attempt_count + 1
                If attempt_count > 9 Then Exit Do
            Loop

            EMReadScreen new_ssn, 11, 7, 42
            new_ssn = replace(new_ssn, " ", "")
            If new_ssn = numb_only_ssn Then ssn_update_success = True
        End If

        If curr_ssn_verif <> left(HH_MEMB_ARRAY(ssn_verif, the_memb), 1)        Then
            ' CHANGES_ARRAY(ssn_verif, the_memb) = curr_ssn_verif
			If curr_ssn_verif = "A" THen CHANGES_ARRAY(ssn_verif, the_memb) = "A - SSN Applied For"
			If curr_ssn_verif = "P" THen CHANGES_ARRAY(ssn_verif, the_memb) = "P - SSN Provided, verif Pending"
			If curr_ssn_verif = "N" THen CHANGES_ARRAY(ssn_verif, the_memb) = "N - SSN Not Provided"
			If curr_ssn_verif = "V" THen CHANGES_ARRAY(ssn_verif, the_memb) = "V - SSN Verified via Interface"
        End If
        If HH_MEMB_ARRAY(ssn_verif, the_memb) = "N - Member Does Not Have SSN" Then CHANGES_ARRAY(ssn_verif, the_memb) = "N - Member Does Not Have SSN"
        If curr_birthdate_verif <> left(HH_MEMB_ARRAY(birthdate_verif, the_memb), 2) Then
            ' CHANGES_ARRAY(birthdate_verif, the_memb) = curr_birthdate_verif
			If curr_birthdate_verif = "BC" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "BC - Birth Certificate"
			If curr_birthdate_verif = "RE" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "RE - Religious Record"
			If curr_birthdate_verif = "DL" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "DL - Drivers License/State ID"
			If curr_birthdate_verif = "DV" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "DV - Divorce Decree"
			If curr_birthdate_verif = "AL" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "AL - Alien Card"
			If curr_birthdate_verif = "DR" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "DR - Doctor Statement"
			If curr_birthdate_verif = "OT" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "OT - Other Document"
			If curr_birthdate_verif = "PV" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "PV - Passport/Visa"
			If curr_birthdate_verif = "NO" Then CHANGES_ARRAY(birthdate_verif, the_memb) = "NO - No Verif Provided"
        End If
        If curr_gender <> left(HH_MEMB_ARRAY(gender, the_memb), 1)              Then
            CHANGES_ARRAY(gender, the_memb) = curr_gender
            If curr_gender = "M" Then CHANGES_ARRAY(gender, the_memb) = "Male"
            If curr_gender = "F" Then CHANGES_ARRAY(gender, the_memb) = "Female"
        End If
        If curr_race <> HH_MEMB_ARRAY(race, the_memb)                           Then CHANGES_ARRAY(race, the_memb) = curr_race
        If curr_spoken_lang <> left(HH_MEMB_ARRAY(spoken_lang, the_memb), 2)    Then CHANGES_ARRAY(spoken_lang, the_memb) = curr_spoken_lang
        If curr_written_lang <> left(HH_MEMB_ARRAY(written_lang, the_memb), 2)  Then CHANGES_ARRAY(written_lang, the_memb) = curr_written_lang
        If curr_interpreter <> left(HH_MEMB_ARRAY(interpreter, the_memb), 1)    Then
            CHANGES_ARRAY(interpreter, the_memb) = curr_interpreter
            If curr_interpreter = "Y" Then CHANGES_ARRAY(interpreter, the_memb) = "Yes"
            If curr_interpreter = "N" Then CHANGES_ARRAY(interpreter, the_memb) = "No"
        End If
        If curr_alias_yn <> left(HH_MEMB_ARRAY(alias_yn, the_memb), 1)          Then CHANGES_ARRAY(alias_yn, the_memb) = curr_alias_yn
        If curr_ethnicity_yn <> left(HH_MEMB_ARRAY(ethnicity_yn, the_memb), 1)  Then CHANGES_ARRAY(ethnicity_yn, the_memb) = curr_ethnicity_yn

        EMReadScreen curr_rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
		EMReadScreen curr_id_verif, 2, 9, 68

        If curr_rel_to_applcnt <> left(HH_MEMB_ARRAY(rel_to_applcnt, the_memb), 2)  Then CHANGES_ARRAY(rel_to_applcnt, the_memb) = curr_rel_to_applcnt
        If curr_id_verif <> left(HH_MEMB_ARRAY(id_verif, the_memb), 2)              Then CHANGES_ARRAY(id_verif, the_memb) = curr_id_verif

        Call navigate_to_MAXIS_screen("STAT", "MEMI")		'===============================================================================================
        EMWriteScreen HH_MEMB_ARRAY(ref_number, the_memb), 20, 76
        transmit

        EMReadScreen curr_marital_status, 1, 7, 40
        EMReadScreen curr_last_grade_completed, 2, 10, 49
        EMReadScreen curr_citizen, 1, 11, 49
        EMReadScreen curr_mn_entry_date, 8, 15, 49
        EMReadScreen curr_former_state, 2, 15, 78

        curr_mn_entry_date = replace(curr_mn_entry_date, " ", "/")
        If curr_mn_entry_date = "__/__/__" Then curr_mn_entry_date = ""
        If curr_former_state = "__" Then curr_former_state = ""

        If curr_marital_status <> left(HH_MEMB_ARRAY(marital_status, the_memb), 1)              Then CHANGES_ARRAY(marital_status, the_memb) = curr_marital_status
        If curr_last_grade_completed <> right(HH_MEMB_ARRAY(last_grade_completed, the_memb), 2) Then CHANGES_ARRAY(last_grade_completed, the_memb) = curr_last_grade_completed
        If curr_citizen <> left(HH_MEMB_ARRAY(citizen, the_memb), 1)                            Then CHANGES_ARRAY(citizen, the_memb) = curr_citizen
        If curr_mn_entry_date <> HH_MEMB_ARRAY(mn_entry_date, the_memb)                         Then CHANGES_ARRAY(mn_entry_date, the_memb) = curr_mn_entry_date
        If curr_former_state <> HH_MEMB_ARRAY(former_state, the_memb)                           Then CHANGES_ARRAY(former_state, the_memb) = curr_former_state

        'THESE ARE NOT EDITABLE
        ' EMReadScreen HH_MEMB_ARRAY(spouse_ref, clt_count), 2, 9, 49
        ' EMReadScreen HH_MEMB_ARRAY(spouse_name, clt_count), 40, 9, 52
        ' EMReadScreen HH_MEMB_ARRAY(other_st_FS_end_date, clt_count), 8, 13, 49
        ' EMReadScreen HH_MEMB_ARRAY(in_mn_12_mo, clt_count), 1, 14, 49
        ' EMReadScreen HH_MEMB_ARRAY(residence_verif, clt_count), 1, 14, 78

    End If
Next

' complete_interview_msg = MsgBox("This interview is now completed and has taken " & interview_time & " minutes." & vbCr & vbCr & "The script will now create your interview notes in a PDF and enter CASE:NOTE(s) as needed.", vbInformation, "Interview Completed")

' script_end_procedure("At this point the script will create a PDF with all of the interview notes to save to ECF, enter a comprehensive CASE:NOTE, and update PROG or REVW with the interview date. Future enhancements will add more actions functionality.")
'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "Form Details - CASE #" & MAXIS_case_number			'Title of the document
objWord.Visible = False 														'The worker should NOT see the docuement
' objWord.Visible = True														'Let the worker see the document
'allow certain workers to see the document
' If user_ID_for_validation = "WFA168" or user_ID_for_validation = "LILE002" Then objWord.Visible = True

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
objSelection.TypeText "DATE OF APPLICATION: " & CAF_datestamp & vbCR
objSelection.TypeText "APPLICATION FORM: " & CAF_form_name & vbCR
objSelection.TypeText "Interview Date: " & interview_date & vbCR
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
If CASH_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", Cash"
If GRH_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", HS/GRH"
If SNAP_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", SNAP"
If EMER_on_CAF_checkbox = checked Then caf_progs = caf_progs & ", EMER"
If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
objSelection.TypeText "PROGRAMS REQUESTED ON FORM: " & caf_progs & vbCr

progs_verbal_request = ""
If cash_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", Cash"
If grh_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", HS/GRH"
If snap_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", SNAP"
If emer_verbal_request = "Yes" Then progs_verbal_request = progs_verbal_request & ", EMER"
If left(progs_verbal_request, 2) = ", " Then progs_verbal_request = right(progs_verbal_request, len(progs_verbal_request)-2)
If progs_verbal_request <> "" Then objSelection.TypeText "PROGRAMS REQUESTED VERBALLY IN INTERVIEW: " & progs_verbal_request & vbCr

objSelection.Font.Size = "11"

If CAF_form = "MNbenefits" Then
    objSelection.TypeText "MN Benefits Appliction Cover Letter Details:" & vbCr
    If trim(additional_application_comments) <> "" Then objSelection.TypeText " - Additional Application Comments: " & additional_application_comments & vbCr
    If trim(additional_application_comments) = "" Then objSelection.TypeText " - No additional application comments." & vbCr
    If trim(additional_income_comments) <> "" Then objSelection.TypeText " - Additional Income Comments: " & additional_income_comments & vbCr
    If trim(additional_income_comments) = "" Then objSelection.TypeText " - No additional income comments." & vbCr
    If trim(cover_letter_interview_notes) <> "" Then objSelection.TypeText " - Interview Notes on Cover Letter Details: " & cover_letter_interview_notes & vbCr
    objSelection.TypeText " - Any income or emergency details from the cover letter can be found in the specific questions further down in this document." & vbCr & vbCr
End If

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

If CHANGES_ARRAY(last_name_const, 0) <> ""      Then objSelection.TypeText "Last Name Changed from " & CHANGES_ARRAY(last_name_const, 0) & " to " & HH_MEMB_ARRAY(last_name_const, 0) & vbCr
If CHANGES_ARRAY(first_name_const, 0) <> ""     Then objSelection.TypeText "First Name Changed from " & CHANGES_ARRAY(first_name_const, 0) & " to " & HH_MEMB_ARRAY(first_name_const, 0) & vbCr
If CHANGES_ARRAY(mid_initial, 0) <> ""          Then objSelection.TypeText "Middle Initial Changed from " & CHANGES_ARRAY(mid_initial, 0) & " to " & HH_MEMB_ARRAY(mid_initial, 0) & vbCr
If CHANGES_ARRAY(date_of_birth, 0) <> ""        Then objSelection.TypeText "Date of Birth Changed from " & CHANGES_ARRAY(date_of_birth, 0) & " to " & HH_MEMB_ARRAY(date_of_birth, 0) & vbCr
If CHANGES_ARRAY(birthdate_verif, 0) <> ""      Then objSelection.TypeText "DoB Verification Changed from " & CHANGES_ARRAY(birthdate_verif, 0) & " to " & HH_MEMB_ARRAY(birthdate_verif, 0) & vbCr
If CHANGES_ARRAY(age, 0) <> ""                  Then objSelection.TypeText "Age Changed from " & CHANGES_ARRAY(age, 0) & " to " & HH_MEMB_ARRAY(age, 0) & vbCr
If CHANGES_ARRAY(ssn, 0) <> ""                  Then objSelection.TypeText "SSN Updated." & vbCr '" from " & CHANGES_ARRAY(ssn, 0) & " to " & HH_MEMB_ARRAY(ssn, 0) & vbCr
If CHANGES_ARRAY(ssn_verif, 0) <> ""            Then objSelection.TypeText "SSN Verification Changed from " & CHANGES_ARRAY(ssn_verif, 0) & " to " & HH_MEMB_ARRAY(ssn_verif, 0) & vbCr
If CHANGES_ARRAY(spoken_lang, 0) <> ""          Then objSelection.TypeText "Spoken Language Changed from " & CHANGES_ARRAY(spoken_lang, 0) & " to " & HH_MEMB_ARRAY(spoken_lang, 0) & vbCr
If CHANGES_ARRAY(written_lang, 0) <> ""         Then objSelection.TypeText "Written Language Changed from " & CHANGES_ARRAY(written_lang, 0) & " to " & HH_MEMB_ARRAY(written_lang, 0) & vbCr
If CHANGES_ARRAY(interpreter, 0) <> ""          Then objSelection.TypeText "Interpreter Needed Changed from " & CHANGES_ARRAY(interpreter, 0) & " to " & HH_MEMB_ARRAY(interpreter, 0) & vbCr
If CHANGES_ARRAY(alias_yn, 0) = "Y"             Then objSelection.TypeText "Alias Name Added: " & HH_MEMB_ARRAY(other_names, 0) & vbCr
If CHANGES_ARRAY(alias_yn, 0) = "N"             Then objSelection.TypeText "Alias Name Removed." & vbCr
' If CHANGES_ARRAY(alias_yn, 0) <> ""             Then objSelection.TypeText "XXXX Changed from " & CHANGES_ARRAY(alias_yn, 0) & " to " & HH_MEMB_ARRAY(alias_yn, 0) & vbCr
If CHANGES_ARRAY(gender, 0) <> ""               Then objSelection.TypeText "Gender Changed from " & CHANGES_ARRAY(gender, 0) & " to " & HH_MEMB_ARRAY(gender, 0) & vbCr
If CHANGES_ARRAY(race, 0) <> ""                 Then objSelection.TypeText "Race Changed from " & CHANGES_ARRAY(race, 0) & " to " & HH_MEMB_ARRAY(race, 0) & vbCr
If CHANGES_ARRAY(ethnicity_yn, 0) <> ""         Then objSelection.TypeText "Ethnicity Changed from " & CHANGES_ARRAY(ethnicity_yn, 0) & " to " & HH_MEMB_ARRAY(ethnicity_yn, 0) & vbCr
If CHANGES_ARRAY(rel_to_applcnt, 0) <> ""       Then objSelection.TypeText "Relationship to Applicant Changed from " & CHANGES_ARRAY(rel_to_applcnt, 0) & " to " & HH_MEMB_ARRAY(rel_to_applcnt, 0) & vbCr
If CHANGES_ARRAY(id_verif, 0) <> ""             Then objSelection.TypeText "ID Verification Changed from " & CHANGES_ARRAY(id_verif, 0) & " to " & HH_MEMB_ARRAY(id_verif, 0) & vbCr
If CHANGES_ARRAY(marital_status, 0) <> ""       Then objSelection.TypeText "Marital Status Changed from " & CHANGES_ARRAY(marital_status, 0) & " to " & HH_MEMB_ARRAY(marital_status, 0) & vbCr
If CHANGES_ARRAY(last_grade_completed, 0) <> "" Then objSelection.TypeText "Last Grade Completed Changed from " & CHANGES_ARRAY(last_grade_completed, 0) & " to " & HH_MEMB_ARRAY(last_grade_completed, 0) & vbCr
If CHANGES_ARRAY(citizen, 0) <> ""              Then objSelection.TypeText "Citizen Changed from " & CHANGES_ARRAY(citizen, 0) & " to " & HH_MEMB_ARRAY(citizen, 0) & vbCr
If CHANGES_ARRAY(mn_entry_date, 0) <> ""        Then objSelection.TypeText "MN Entry Date Changed from " & CHANGES_ARRAY(mn_entry_date, 0) & " to " & HH_MEMB_ARRAY(mn_entry_date, 0) & vbCr
If CHANGES_ARRAY(former_state, 0) <> ""         Then objSelection.TypeText "Former State Changed from " & CHANGES_ARRAY(former_state, 0) & " to " & HH_MEMB_ARRAY(former_state, 0) & vbCr

objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, 0) & vbCR
objSelection.TypeText chr(9) & "Identity: " & HH_MEMB_ARRAY(id_verif, 0) & vbCr
If HH_MEMB_ARRAY(intend_to_reside_in_mn, 0) <> "" Then objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, 0) & vbCr
If HH_MEMB_ARRAY(imig_status, 0) <> "" Then objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, 0) & vbCr
If HH_MEMB_ARRAY(clt_has_sponsor, 0) <> "" and HH_MEMB_ARRAY(clt_has_sponsor, 0) <> "?" Then objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, 0) & vbCr
objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, 0) & vbCr
If HH_MEMB_ARRAY(client_verification_details, 0) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, 0) & vbCr


objSelection.TypeText vbCr & "Contact Info" & vbCR
If send_text = "Yes" Then objSelection.TypeText "Send Updates Via Text Message: Yes" & vbCR
If send_text = "No" Then objSelection.TypeText "Send Updates Via Text Message: No" & vbCR
If send_email = "Yes" Then objSelection.TypeText "Send Updates Via E-Mail: Yes" & vbCR
If send_email = "No" Then objSelection.TypeText "Send Updates Via E-Mail: No" & vbCR

objSelection.TypeText vbCr & "Housing Support Info" & vbCR
objSelection.TypeText "-Currently residing in a licensed facility: " & licensed_facility & vbCR
objSelection.TypeText "-Residence provides meals? " & meal_provided & vbCR
objSelection.TypeText "-Name/Phone number of residence: " & residence_name_phone & vbCR
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.TypeText "Household Lives in " & resi_addr_county & " County" & vbCR
If disc_out_of_county = "RESOLVED" Then objSelection.TypeText "- Household reported living Out of Hennepin County - Case Needs Transfer - additional interview conversation: " & disc_out_of_county_confirmation & vbCr

If trim(all_members_listed_notes) <> "" Then objSelection.TypeText vbCr & "HH Comp Notes: " & all_members_listed_notes & vbCR
If all_members_in_MN_yn <> "" Then
    objSelection.TypeText "ALL HH members Intend to reside in MN: " & all_members_in_MN_yn & vbCR
    If trim(all_members_in_MN_notes) <> "" Then objSelection.TypeText " - Notes: " & all_members_in_MN_notes & vbCR
End If
If anyone_pregnant_yn <> "" Then
    objSelection.TypeText "Anyone Pregnant: " & anyone_pregnant_yn & vbCR
    If trim(anyone_pregnant_notes) <> "" Then objSelection.TypeText " - Notes: " & anyone_pregnant_notes & vbCR
End If
If anyone_served_yn <> "" Then
    objSelection.TypeText "Anyone Served in Military: " & anyone_served_yn & vbCR
    If trim(anyone_served_notes) <> "" Then objSelection.TypeText " - Notes: " & anyone_served_notes & vbCR
End If


objSelection.TypeText "LIVING SITUATION: " & living_situation & vbCR

If disc_homeless_no_mail_addr = "RESOLVED" Then objSelection.TypeText "- Household Experiencing Housing Insecurity - MAIL is Primary Communication of Agency Requests and Actions - additional interview conversation: " & disc_homeless_confirmation & vbCr
If disc_no_phone_number = "RESOLVED" Then objSelection.TypeText "- No Phone Number was Provided - additional interview conversation: " & disc_phone_confirmation & vbCr

'Now we have a dynamic number of tables
'each table has to be defined with its index so we need to have a variable to increment
table_count = 3			'table index variable

If expedited_screening_on_form = True Then

	' objSelection.Font.Bold = TRUE
	objSelection.TypeText "EXPEDITED QUESTIONS from the Form"
	Set objRange = objSelection.Range					'range is needed to create tables
	objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
	set objEXPTable = objDoc.Tables(table_count)		'Creates the table with the specific index'
	table_count = table_count + 1

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

End If

objSelection.TypeParagraph()						'adds a line between the table and the next information

If expedited_determination_needed = True Then
	objSelection.Font.Bold = TRUE
	objSelection.TypeText "EXPEDITED Interview Answers:" & vbCr
	objSelection.Font.Bold = FALSE
	If is_elig_XFS = True Then
		objSelection.TypeText "Based on income information this case APPEARS ELIGIBLE FOR EXPEDITED SNAP." & vbCr
	Else
		objSelection.TypeText "This case does not appear eligible for expedited SNAP based on the income information." & vbCr
	End If
	objSelection.TypeText chr(9) & "Income in the month of application: " & determined_income & vbCr
	objSelection.TypeText chr(9) & "Assets in the month of application: " & determined_assets & vbCr
	objSelection.TypeText chr(9) & "Expenses in the month of application: " & calculated_expenses & vbCr
	objSelection.TypeText chr(9) & chr(9) & "Housing expense in the month of application: " & determined_shel & vbCr
	objSelection.TypeText chr(9) & chr(9) & "Utilities in the month of application: " & determined_utilities & vbCr
	If trim(exp_det_notes) <> "" Then objSelection.TypeText chr(9) & "Additional Notes: " & exp_det_notes & vbCr
	If trim(delay_explanation) <> "" Then
		objSelection.TypeText chr(9) & "Expedited Approval must be delayed:" & vbCr
		line_start = chr(9) & chr(9) & "Detail: "
		counter = 1
		If InStr(delay_explanation, ";") = 0 Then
			objSelection.TypeText line_start & counter & ". " & delay_explanation & vbCr
			line_start = chr(9) & chr(9) & chr(9)
		Else
			delay_explain_array = Split(delay_explanation, ";")
			For each delay_reason in delay_explain_array
				delay_reason = trim(delay_reason)
				objSelection.TypeText line_start & counter & ". " & delay_reason & vbCr
				line_start = chr(9) & chr(9) & chr(9)
				counter = counter + 1
			Next
		End If
	End If

	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
	objSelection.TypeParagraph()						'adds a line between the table and the next information
End If

objSelection.Font.Bold = TRUE
objSelection.TypeText "Interview Answers:" & vbCr
objSelection.Font.Bold = FALSE

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

            If NOT HH_MEMB_ARRAY(pers_in_maxis, each_member) Then objSelection.TypeText "Person needs to be added to MAXIS." & vbCr

            If HH_MEMB_ARRAY(pers_in_maxis, each_member) Then
                If CHANGES_ARRAY(last_name_const, each_member) <> ""      Then objSelection.TypeText "Last Name Changed from " & CHANGES_ARRAY(last_name_const, each_member) & " to " & HH_MEMB_ARRAY(last_name_const, each_member) & vbCr
                If CHANGES_ARRAY(first_name_const, each_member) <> ""     Then objSelection.TypeText "First Name Changed from " & CHANGES_ARRAY(first_name_const, each_member) & " to " & HH_MEMB_ARRAY(first_name_const, each_member) & vbCr
                If CHANGES_ARRAY(mid_initial, each_member) <> ""          Then objSelection.TypeText "Middle Initial Changed from " & CHANGES_ARRAY(mid_initial, each_member) & " to " & HH_MEMB_ARRAY(mid_initial, each_member) & vbCr
                If CHANGES_ARRAY(date_of_birth, each_member) <> ""        Then objSelection.TypeText "Date of Birth Changed from " & CHANGES_ARRAY(date_of_birth, each_member) & " to " & HH_MEMB_ARRAY(date_of_birth, each_member) & vbCr
                If CHANGES_ARRAY(birthdate_verif, each_member) <> ""      Then objSelection.TypeText "DoB Verification Changed from " & CHANGES_ARRAY(birthdate_verif, each_member) & " to " & HH_MEMB_ARRAY(birthdate_verif, each_member) & vbCr
                If CHANGES_ARRAY(age, each_member) <> ""                  Then objSelection.TypeText "Age Changed from " & CHANGES_ARRAY(age, each_member) & " to " & HH_MEMB_ARRAY(age, each_member) & vbCr
                If CHANGES_ARRAY(ssn, each_member) <> ""                  Then objSelection.TypeText "SSN Updated." & vbCr '" from " & CHANGES_ARRAY(ssn, each_member) & " to " & HH_MEMB_ARRAY(ssn, each_member) & vbCr
                If CHANGES_ARRAY(ssn_verif, each_member) <> ""            Then objSelection.TypeText "SSN Verification Changed from " & CHANGES_ARRAY(ssn_verif, each_member) & " to " & HH_MEMB_ARRAY(ssn_verif, each_member) & vbCr
                If CHANGES_ARRAY(spoken_lang, each_member) <> ""          Then objSelection.TypeText "Spoken Language Changed from " & CHANGES_ARRAY(spoken_lang, each_member) & " to " & HH_MEMB_ARRAY(spoken_lang, each_member) & vbCr
                If CHANGES_ARRAY(written_lang, each_member) <> ""         Then objSelection.TypeText "Written Language Changed from " & CHANGES_ARRAY(written_lang, each_member) & " to " & HH_MEMB_ARRAY(written_lang, each_member) & vbCr
                If CHANGES_ARRAY(interpreter, each_member) <> ""          Then objSelection.TypeText "Interpreter Needed Changed from " & CHANGES_ARRAY(interpreter, each_member) & " to " & HH_MEMB_ARRAY(interpreter, each_member) & vbCr
                If CHANGES_ARRAY(alias_yn, each_member) = "Y"             Then objSelection.TypeText "Alias Name Added: " & HH_MEMB_ARRAY(other_names, each_member) & vbCr
                If CHANGES_ARRAY(alias_yn, each_member) = "N"             Then objSelection.TypeText "Alias Name Removed." & vbCr
                ' If CHANGES_ARRAY(alias_yn, each_member) <> ""             Then objSelection.TypeText "XXXX Changed from " & CHANGES_ARRAY(alias_yn, each_member) & " to " & HH_MEMB_ARRAY(alias_yn, each_member) & vbCr
                If CHANGES_ARRAY(gender, each_member) <> ""               Then objSelection.TypeText "Gender Changed from " & CHANGES_ARRAY(gender, each_member) & " to " & HH_MEMB_ARRAY(gender, each_member) & vbCr
                If CHANGES_ARRAY(race, each_member) <> ""                 Then objSelection.TypeText "Race Changed from " & CHANGES_ARRAY(race, each_member) & " to " & HH_MEMB_ARRAY(race, each_member) & vbCr
                If CHANGES_ARRAY(ethnicity_yn, each_member) <> ""         Then objSelection.TypeText "Ethnicity Changed from " & CHANGES_ARRAY(ethnicity_yn, each_member) & " to " & HH_MEMB_ARRAY(ethnicity_yn, each_member) & vbCr
                If CHANGES_ARRAY(rel_to_applcnt, each_member) <> ""       Then objSelection.TypeText "Relationship to Applicant Changed from " & CHANGES_ARRAY(rel_to_applcnt, each_member) & " to " & HH_MEMB_ARRAY(rel_to_applcnt, each_member) & vbCr
                If CHANGES_ARRAY(id_verif, each_member) <> ""             Then objSelection.TypeText "ID Verification Changed from " & CHANGES_ARRAY(id_verif, each_member) & " to " & HH_MEMB_ARRAY(id_verif, each_member) & vbCr
                If CHANGES_ARRAY(marital_status, each_member) <> ""       Then objSelection.TypeText "Marital Status Changed from " & CHANGES_ARRAY(marital_status, each_member) & " to " & HH_MEMB_ARRAY(marital_status, each_member) & vbCr
                If CHANGES_ARRAY(last_grade_completed, each_member) <> "" Then objSelection.TypeText "Last Grade Completed Changed from " & CHANGES_ARRAY(last_grade_completed, each_member) & " to " & HH_MEMB_ARRAY(last_grade_completed, each_member) & vbCr
                If CHANGES_ARRAY(citizen, each_member) <> ""              Then objSelection.TypeText "Citizen Changed from " & CHANGES_ARRAY(citizen, each_member) & " to " & HH_MEMB_ARRAY(citizen, each_member) & vbCr
                If CHANGES_ARRAY(mn_entry_date, each_member) <> ""        Then objSelection.TypeText "MN Entry Date Changed from " & CHANGES_ARRAY(mn_entry_date, each_member) & " to " & HH_MEMB_ARRAY(mn_entry_date, each_member) & vbCr
                If CHANGES_ARRAY(former_state, each_member) <> ""         Then objSelection.TypeText "Former State Changed from " & CHANGES_ARRAY(former_state, each_member) & " to " & HH_MEMB_ARRAY(former_state, each_member) & vbCr
            End If

    		objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, each_member) & vbCR
    		' objSelection.Font.Bold = TRUE
    		' objSelection.TypeText "AGENCY USE:" & vbCr
    		' objSelection.Font.Bold = FALSE
    		objSelection.TypeText chr(9) & "Identity: " & HH_MEMB_ARRAY(id_verif, each_member) & vbCr
    		If HH_MEMB_ARRAY(intend_to_reside_in_mn, each_member) <> "" Then objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, each_member) & vbCr
    		If HH_MEMB_ARRAY(imig_status, each_member) <> "" Then objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, each_member) & vbCr
    		If HH_MEMB_ARRAY(clt_has_sponsor, each_member) <> "" and HH_MEMB_ARRAY(clt_has_sponsor, each_member) <> "?" Then objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, each_member) & vbCr
    		objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, each_member) & vbCr
    		If HH_MEMB_ARRAY(client_verification_details, each_member) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, each_member) & vbCr

    		array_counters = array_counters + 1
        End If
	Next
Else
	objSelection.TypeText "THERE ARE NO OTHER PEOPLE TO BE LISTED ON THIS APPLICATION" & vbCr
	ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
End If


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

For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
	FORM_QUESTION_ARRAY(each_question).add_to_wif()
Next

objSelection.TypeText "QUALIFYING QUESTIONS" & vbCr

objSelection.TypeText "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the Form?" & vbCr
objSelection.TypeText chr(9) & qual_question_one & vbCr
If trim(qual_memb_one) <> "" AND qual_memb_one <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_one & vbCr
objSelection.TypeText "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?" & vbCr
objSelection.TypeText chr(9) & qual_question_two & vbCr
If trim(qual_memb_two) <> "" AND qual_memb_two <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_two & vbCr
objSelection.TypeText "Is anyone in your household hiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?" & vbCr
objSelection.TypeText chr(9) & qual_question_three & vbCr
If trim(qual_memb_three) <> "" AND qual_memb_three <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_three & vbCr
objSelection.TypeText "Has anyone in your household been convicted of a drug felony in the past 10 years?" & vbCr
objSelection.TypeText chr(9) & qual_question_four & vbCr
If trim(qual_memb_four) <> "" AND qual_memb_four <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_four & vbCr
objSelection.TypeText "Is anyone in your household currently violating a condition of parole, probation or supervised release?" & vbCr
objSelection.TypeText chr(9) & qual_question_five & vbCr
If trim(qual_memb_five) <> "" AND qual_memb_five <> "Select or Type" Then objSelection.TypeText chr(9) & qual_memb_five & vbCr

resident_emergency_yn = trim(resident_emergency_yn)
If resident_emergency_yn <> "" or (trim(emergency_type) <> "" and trim(emergency_type) <> "Select or Type") or trim(emergency_discussion) <> "" or trim(emergency_amount) <> "" or trim(emergency_deadline) <> "" Then
	objSelection.TypeText "EMERGENCY QUESTIONS" & vbCr
	If resident_emergency_yn <> "" Then
		objSelection.TypeText "Is the resident experiencing an emergency?" & vbCr
		objSelection.TypeText chr(9) & resident_emergency_yn & vbCr
	End If
	If trim(emergency_type) <> "" and trim(emergency_type) <> "Select or Type" Then
		objSelection.TypeText "What emergency is the resident is experiencing?" & vbCr
		objSelection.TypeText chr(9) & emergency_type & vbCr
	End If
	If trim(emergency_discussion) <> "" Then
		objSelection.TypeText "Discussion of emergency with resident:" & vbCr
		objSelection.TypeText chr(9) & emergency_discussion & vbCr
	End If
	If trim(emergency_amount) <> ""  Then
		objSelection.TypeText "What amount is needed to resolve the emergency?" & vbCr
		objSelection.TypeText chr(9) & emergency_amount & vbCr
	End If
	If trim(emergency_deadline) <> "" Then
		objSelection.TypeText "What is the deadline to resolve the emergency?" & vbCr
		objSelection.TypeText chr(9) & emergency_deadline & vbCr
	End If
Else
	objSelection.TypeText "No Emergency Details Recorded" & vbCr
End If

objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

objSelection.TypeText "Signatures:" & vbCr
objSelection.Font.Size = "12"

If signature_detail <> "Not Required" AND signature_detail <> "Blank" Then
	objSelection.TypeText "Signature of Primary Adult: " & signature_person & ", " & signature_detail & vbCr
ElseIf signature_detail = "Blank" Then
	objSelection.TypeText "Signature of Primary Adult is blank." & vbCr
End If
If second_signature_detail <> "Not Required" AND second_signature_detail <> "Blank" Then
	objSelection.TypeText "Signature of Secondary Adult: " & second_signature_person & ", " & second_signature_detail & vbCr
ElseIf second_signature_detail = "Blank" Then
	objSelection.TypeText "Signature of Secondary Adult is blank." & vbCr
End If
If signature_detail = "Accepted Verbally" or second_signature_detail = "Accepted Verbally" Then
	objSelection.TypeText "Verbal Signature Accepted during interview on " & verbal_sig_date & " at " & verbal_sig_time & "." & vbCr
	objSelection.TypeText "Resident Phone Number: " & verbal_sig_phone_number & vbCr
End If
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

	If arep_on_CAF_checkbox = checked Then objSelection.TypeText "This AREP information was entered on the Form." & vbCR
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
		objSelection.TypeText "The Housing Expense information within the Form do not appear to Match" & vbCr
		objSelection.TypeText "  - Page 1 Housing Expense: " & exp_q_3_rent_this_month & vbCr
		objSelection.TypeText "  - Question on Housing Expense: " & rent_summary & vbCr
		objSelection.TypeText "  - Resolution: " & disc_rent_amounts_confirmation & vbCr
	End If
	If disc_utility_amounts = "RESOLVED" Then
		objSelection.TypeText "The Utility Expense information within the Form do not appear to Match" & vbCr
		objSelection.TypeText "  - Page 1 Utility Expense: " & disc_utility_caf_1_summary & vbCr
		objSelection.TypeText "  - Question on Utility Expense: " & utility_summary & vbCr
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
' MsgBox pdf_doc_path
If developer_mode = True Then pdf_doc_path = t_drive & "\Eligibility Support\Assignments\Interview Notes for ECF\Archive\TRAINING REGION Interviews - NOT for ECF\Interview - " & MAXIS_case_number & " on " & file_safe_date & ".pdf"

'Now we save the document.
'MS Word allows us to save directly as a PDF instead of a DOC.
'the file path must be PDF
'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
objDoc.SaveAs pdf_doc_path, 17
STATS_manualtime = STATS_manualtime + 60

'This looks to see if the PDF file has been correctly saved. If it has the file will exists in the pdf file path
If objFSO.FileExists(pdf_doc_path) = TRUE Then
	'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
	objDoc.Close wdDoNotSaveChanges
	objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)

	'Now we MEMO'
	STATS_manualtime = STATS_manualtime + 195
	Call start_a_new_spec_memo(memo_opened, False, "N", "N", "N", other_name, other_street, other_city, other_state, other_zip, False)

	If memo_opened = True Then
		CALL write_variable_in_SPEC_MEMO("You have completed your interview on " & interview_date)
		CALL write_variable_in_SPEC_MEMO("This is for the " & CAF_form_name & " you submitted.")
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
	End If
	If arep_authorization <> "DO NOT AUTHORIZE AN AREP" Then
		If arep_exists = True Then
			If arep_action = "Yes - keep this AREP" OR CAF_arep_action = "Yes - add to MAXIS" OR (arep_authorization <> "Select One..." AND arep_authorization <> "") Then
				Call start_a_new_spec_memo(memo_opened, False, "N", "N", "N", other_name, other_street, other_city, other_state, other_zip, False)
				If memo_opened = True Then
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
				End If

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
				PF3
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
				Set objTextStream = nothing

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
					For each delay_explain in delay_explain_array
						delay_explain = trim(delay_explain)
						Call write_variable_with_indent_in_CASE_NOTE(counter & ". " & delay_explain)
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
		Call write_bullet_and_variable_in_CASE_NOTE("Notes", exp_det_notes)

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

	'setting the end message
	end_msg = "Success! The information you have provided about the interview and all of the notes have been saved in a PDF. This PDF will be uploaded to ECF by the ES Support Team for Case # " & MAXIS_case_number & " and will remain in the CASE RECORD. CASE:NOTES have also been entered with the full interview detail."
	o2Exec.Terminate()
	'Now we ask if the worker would like the PDF to be opened by the script before the script closes
	'This is helpful because they may not be familiar with where these are saved and they could work from the PDF to process the reVw
	reopen_pdf_doc_msg = MsgBox("The information gathered in the interview has been saved as a PDF and will be added to ECF as a separate 'Interview Notes' document." & vbCr & vbCr & "This document will take the place of the INTERVIEW ANNOTATIONS on the form in ECF, as long as you have entered all interview notes to the script." & vbCr & "Agency Signature is not required on the application form." & vbCr & vbCr & "Would you like the PDF Document opened to process/review?", vbQuestion + vbSystemModal + vbYesNo, "Open PDF Doc?")
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

end_msg = end_msg & vbCr & vbCr & "Form received: " & CAF_form_name
end_msg = end_msg & vbCr & vbCr & "The documment created for the ECF Case File can serve in place of any annotations as long as you entered all of your interview notes into the script. If you have entered all of the interview notes for this interview, there is no need to annotate the application form in ECF."
end_msg = end_msg & vbCr & vbCr & "Hennepin County does not require an Agency Signature to be added to the application form. Details can be found in the HSR Manual: https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/Applications.aspx (Search: Applications)."
With (CreateObject("Scripting.FileSystemObject"))
	.DeleteFile(intvw_done_msg_file)
End With

If run_by_interview_team = True and developer_mode = False Then
	'creates an XML File with details of the the interview
	Set xmlTracDoc = CreateObject("Microsoft.XMLDOM")
	xmlTracPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_details_" & MAXIS_case_number & "_on_" & replace(replace(interview_date, "/", "_")," ", "_") & ".xml"

	xmlTracDoc.async = False

	Set root = xmlTracDoc.createElement("interview")
	xmlTracDoc.appendChild root

	Set element = xmlTracDoc.createElement("ScriptRunDate")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(date)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ScriptRunTime")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(time)
	element.appendChild info

	Set element = xmlTracDoc.createElement("WorkerName")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(worker_name)
	element.appendChild info

	Set element = xmlTracDoc.createElement("WindowsUserID")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(windows_user_ID)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CaseNumber")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(MAXIS_case_number)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CaseBasket")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(case_pw)
	element.appendChild info

	Set element = xmlTracDoc.createElement("DHSFormNumber")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(CAF_form_number)
	element.appendChild info

	Set element = xmlTracDoc.createElement("DHSFormName")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(CAF_form_name)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewDate")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(interview_date)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CaseActive")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(case_active)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CasePending")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(case_pending)
	element.appendChild info

		Set element = xmlTracDoc.createElement("InterviewPerson")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(who_are_we_completing_the_interview_with)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewMethod")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(how_are_we_completing_the_interview)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewLength")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(length_of_interview)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewInterpreter")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(interpreter_information)
	element.appendChild info

	Set element = xmlTracDoc.createElement("InterviewLanguage")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(interpreter_language)
	element.appendChild info

	Set element = xmlTracDoc.createElement("FormInfo")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(CAF_form)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CAFDateStamp")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(CAF_datestamp)
	element.appendChild info

	Set element = xmlTracDoc.createElement("SNAPStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(snap_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(grh_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("MFIPStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(mfip_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("DWPStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(dwp_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GAStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(ga_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("MSAStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(msa_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("EMERStatus")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(emer_status)
	element.appendChild info

	Set element = xmlTracDoc.createElement("UnspecifiedCASHPending")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(unknown_cash_pending)
	element.appendChild info

	Set element = xmlTracDoc.createElement("SNAPClosedPastThirtyDays")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(snap_closed_in_past_30_days)
	element.appendChild info

	Set element = xmlTracDoc.createElement("SNAPClosedPastFourMonths")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(snap_closed_in_past_4_months)
	element.appendChild info

	Set element = xmlTracDoc.createElement("FSDateClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(FS_date_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("FSReasonClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(FS_reason_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHClosedPastThirtyDays")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(grh_closed_in_past_30_days)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHClosedPastFourMonths")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(grh_closed_in_past_4_months)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHDateClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(GRH_date_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHReasonClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(GRH_reason_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHOneClosedPastThirtyDays")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash1_closed_in_past_30_days)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHOneClosedPastFourMonths")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash1_closed_in_past_4_months)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHOneRecentlyClosedProgram")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash1_recently_closed_program)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHOneDateClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash1_date_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHOneClosedReason")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash1_closed_reason)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHTwoClosedPastThirtyDays")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash2_closed_in_past_30_days)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHTwoClosedPastFourMonths")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash2_closed_in_past_4_months)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHTwoRecentlyClosedProgram")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash2_recently_closed_program)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHTwoDateClosed")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash2_date_closed)
	element.appendChild info

	Set element = xmlTracDoc.createElement("CASHTwoClosedReason")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash2_closed_reason)
	element.appendChild info

	Set element = xmlTracDoc.createElement("IssuedDate")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(issued_date)
	element.appendChild info

	Set element = xmlTracDoc.createElement("IssuedProg")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(issued_prog)
	element.appendChild info


	Set element = xmlTracDoc.createElement("CASHRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(cash_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("GRHRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(grh_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("SNAPRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(snap_request)
	element.appendChild info

	Set element = xmlTracDoc.createElement("EMERRequest")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(emer_request)
	element.appendChild info

	If cash_request = True Then
		Set element = xmlTracDoc.createElement("CASHProcess")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(the_process_for_cash)
		element.appendChild info

		Set element = xmlTracDoc.createElement("TypeOfCASH")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(type_of_cash)
		element.appendChild info

		If the_process_for_cash = "Renewal" Then
			Set element = xmlTracDoc.createElement("CASHRenewalMonth")
			root.appendChild element
			Set info = xmlTracDoc.createTextNode(next_cash_revw_mo & "/" & next_cash_revw_yr)
			element.appendChild info
		End If
	End If

	If snap_request = True Then
		Set element = xmlTracDoc.createElement("SNAPProcess")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(the_process_for_snap)
		element.appendChild info
		If the_process_for_snap = "Renewal" Then
			Set element = xmlTracDoc.createElement("CASHRenewalMonth")
			root.appendChild element
			Set info = xmlTracDoc.createTextNode(next_snap_revw_mo & "/" & next_snap_revw_yr)
			element.appendChild info
		End If
	End If

	If emer_request = True Then
		Set element = xmlTracDoc.createElement("EMERProcess")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(the_process_for_emer)
		element.appendChild info
	End If

	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		Set element = xmlTracDoc.createElement("member")
		root.appendChild element

		Set element = xmlTracDoc.createElement("ReferenceNumber")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(ref_number, the_members))
		element.appendChild info

		Set element = xmlTracDoc.createElement("LastName")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(last_name_const, the_members))
		element.appendChild info

		Set element = xmlTracDoc.createElement("FirstName")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(first_name_const, the_members))
		element.appendChild info

		Set element = xmlTracDoc.createElement("Age")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(age, the_members))
		element.appendChild info

		Set element = xmlTracDoc.createElement("RelationshipToApplicant")
		root.appendChild element
		Set info = xmlTracDoc.createTextNode(HH_MEMB_ARRAY(rel_to_applcnt, the_members))
		element.appendChild info

		If HH_MEMB_ARRAY(memb_is_caregiver, the_members) = True Then
			Set element = xmlTracDoc.createElement("MFIPOrientation")
			root.appendChild element
			If HH_MEMB_ARRAY(orientation_needed_const, the_members) = True and HH_MEMB_ARRAY(orientation_done_const, the_members) = False and HH_MEMB_ARRAY(orientation_exempt_const, the_members) = False Then
				Set info = xmlTracDoc.createTextNode("Incomplete")
				element.appendChild info
			ElseIf  HH_MEMB_ARRAY(orientation_needed_const, the_members) = False Then
				Set info = xmlTracDoc.createTextNode("Not Needed")
				element.appendChild info
			ElseIf HH_MEMB_ARRAY(orientation_needed_const, the_members) = True and HH_MEMB_ARRAY(orientation_done_const, the_members) = True Then
				Set info = xmlTracDoc.createTextNode("Completed")
				element.appendChild info
			ElseIf HH_MEMB_ARRAY(orientation_needed_const, the_members) = True and HH_MEMB_ARRAY(orientation_exempt_const, the_members) = True Then
				Set info = xmlTracDoc.createTextNode("Exempt")
				element.appendChild info
			End If
		End If
	Next

	Set element = xmlTracDoc.createElement("eDRSMatchFound")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(edrs_match_found)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpeditedScreening")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(expedited_screening)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpeditedDetermination")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(is_elig_XFS)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpDetIncome")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(exp_det_income)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpDetAssets")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(exp_det_assets)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpDetShelter")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(exp_det_housing)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpDetUtilities")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(exp_det_utilities)
	element.appendChild info

	Set element = xmlTracDoc.createElement("ExpDetNotes")
	root.appendChild element
	Set info = xmlTracDoc.createTextNode(exp_det_notes)
	element.appendChild info

	xmlTracDoc.save(xmlTracPath)

	Set xml = CreateObject("Msxml2.DOMDocument")
	Set xsl = CreateObject("Msxml2.DOMDocument")

	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	txt = Replace(fso.OpenTextFile(xmlTracPath).ReadAll, "><", ">" & vbCrLf & "<")
	stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
	"<xsl:output method=""xml"" indent=""yes""/>" & _
	"<xsl:template match=""/"">" & _
	"<xsl:copy-of select="".""/>" & _
	"</xsl:template>" & _
	"</xsl:stylesheet>"

	xsl.loadXML stylesheet
	xml.loadXML txt

	xml.transformNode xsl

	xml.Save xmlTracPath

	xmlSavePath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_started_" & MAXIS_case_number & ".xml"
	If ObjFSO.FileExists(xmlSavePath) Then objFSO.DeleteFile(xmlSavePath)
End If

STATS_manualtime = STATS_manualtime + (timer - start_time + add_to_time)
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