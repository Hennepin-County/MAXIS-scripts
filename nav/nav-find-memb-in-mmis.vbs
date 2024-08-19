'STATS GATHERING=============================================================================================================
name_of_script = "NAV - FIND MEMB IN MMIS.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 0               'sets the stats counter at one
STATS_manualtime = 120            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "P"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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
'990656
'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("08/13/2024", "Initial version.", "Casey Love, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone

CALL check_for_MAXIS(True)										'Checks to see if in MAXIS
CALL MAXIS_case_number_finder(MAXIS_case_number)    			'Grabs the MAXIS case number automatically
Call back_to_SELF                                               'starting at the SELF panel
EMReadScreen MX_environment, 13, 22, 48                         'seeing which MX environment we are in
MX_environment = trim(MX_environment)
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")        	'Going to MMIS'
Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS

MAXIS_footer_month = CM_mo									'Directly assigns a footer month based on the current month
MAXIS_footer_year = CM_yr

Dialog1 = "" 'blanking out dialog name
BeginDialog Dialog1, 0, 0, 191, 80, "NAV - Find MEMB in MMIS Case Number Dialog"
  EditBox 60, 10, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 135, 40, 50, 15
    CancelButton 135, 60, 50, 15
    PushButton 135, 5, 50, 15, "Instructions", script_instructions_btn
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 40, 120, 35, "Script will allow you to select a single member and find any open Health Care Span in MMIS for the current month."
EndDialog

'Shows dialog ----------------------------------
DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_without_confirmation
        'Add in all of your mandatory field handling from your dialog here.
        Call validate_MAXIS_case_number(err_msg, "*") ' IF NEEDED
		If ButtonPressed = script_instructions_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/NAVIGATION/NAVIGATION%20SCRIPTS.docx"	'copy the instructions URL here
			err_msg = "LOOP"
		End If
        IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

'Reset to SELF to check the MAXIS region
'This is also helpful to ensure we are not starting in a CASE/NOTE or something
Call back_to_SELF
Call clear_line_of_text(18, 43) 					'clear and rewrite the CASE Number. This is optional but can help the worker not to lose the case number.
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

'PRIV Handling
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
If is_this_PRIV = True then script_end_procedure("This case is privileged and you do not have access to it. The script will now end.")

'Out of County Handling
'There are a few reasons to allow a script to run on an out of county case - so review if this is needed.
EMReadScreen pw_county_code, 2, 21, 16
If pw_county_code <> "27" Then script_end_procedure("This case is not in Hennepin County and cannot be updated. The script will now end.")

Do
	Call HH_member_custom_dialog(HH_member_array)
	If UBound(HH_member_array) > 0 Then MsgBox "* * * NOTICE * * * " & vbCr & vbCr & "Search can only be completed for 1 person, please check only one person to search."
Loop until UBound(HH_member_array) = 0

MEMB_ref_numb = HH_member_array(0)
Call find_mmis_pmis_for_memb(MEMB_ref_numb, MEMB_ssn, MEMB_pmi, MEMB_last_name, MEMB_first_name, MEMB_dob, MEMB_PMI_ARRAY)

function find_mmis_pmis_for_memb(MEMB_ref_numb, MEMB_ssn, MEMB_pmi, MEMB_last_name, MEMB_first_name, MEMB_dob, MEMB_PMI_ARRAY)
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	Call write_value_and_transmit(MEMB_ref_numb, 20, 76)

	EMReadScreen MEMB_pmi, 8, 4, 46
	EMReadScreen MEMB_last_name, 25, 6, 30
	EMReadScreen MEMB_first_name, 12, 6, 63
	EMReadScreen MEMB_middle_initial, 1, 6, 79
	EMReadScreen MEMB_ssn, 11, 7, 42
	EMReadScreen MEMB_dob, 10, 8, 42

	MEMB_last_name = replace(MEMB_last_name, "_", "")
	MEMB_first_name = replace(MEMB_first_name, "_", "")
	MEMB_middle_initial = replace(MEMB_middle_initial, "_", "")
	MEMB_dob = replace(MEMB_dob, " ", "/")
	MEMB_pmi = trim(MEMB_pmi)
	If MEMB_ssn = "___ __ ____" Then MEMB_ssn = ""

	'need to get to ground zero
	Call back_to_SELF
	Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")      'Going to MMIS'
	EMReadScreen check_in_MMIS, 18, 1, 7

	If check_in_MMIS = "SESSION TERMINATED" Then
		EMWriteScreen "MW00",1, 2
		transmit
		transmit

		EMWriteScreen "X", 8, 3
		transmit
	End If

	MMIS_PMI_Number = right("00000000" & MEMB_pmi, 8)    'making this 8 charactes because MMIS
	MMIS_SSN = replace(MEMB_ssn, " ", "")
	MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)

	PMIs_String = " " & MMIS_PMI_Number & " "

	For search_count = 1 to 2
		Call get_to_RKEY
		Call clear_line_of_text(4, 19)
		Call clear_line_of_text(5, 19)
		Call clear_line_of_text(6, 19)
		Call clear_line_of_text(6, 48)
		Call clear_line_of_text(6, 69)
		Call clear_line_of_text(7, 19)
		Call clear_line_of_text(7, 19)

		EMWriteScreen "I", 2, 19                                                    'read only
		Select Case search_count
		Case 1
			EMWriteScreen MMIS_SSN, 5, 19
		Case 2
			EMWriteScreen MEMB_last_name, 6, 19
			EMWriteScreen left(MEMB_first_name, 1), 6, 48
			EMWriteScreen MEMB_dob, 7, 19
		End Select
		transmit

		EMReadScreen panel_name, 4, 1, 51
		EMReadScreen panel_name_2, 4, 1, 52

		If panel_name = "RSUM" Then
			EMReadScreen relg_PMI, 8, 2, 2
			If InStr(PMIs_String, relg_PMI) = 0 Then PMIs_String = PMIs_String & relg_PMI & " "
		ElseIf panel_name_2 = "RSEL" Then
			rsel_row = 7
			Do
				capture_pmi = True
				If search_count = 1 Then
					EMReadScreen check_ssn, 9, rsel_row, 48
					If check_ssn <> MMIS_SSN Then capture_pmi = False
				ElseIf search_count = 2 Then
					EMReadScreen check_ssn, 9, rsel_row, 48
					EMReadScreen check_first_name, 13, rsel_row, 33
					check_first_name = trim(check_first_name)
					If check_ssn <> MMIS_SSN and check_first_name <> MEMB_first_name Then capture_pmi = False
				End If

				If capture_pmi = True Then
					EMReadScreen rsel_pmi, 8, rsel_row, 4
					' MsgBox "rsel_pmi - " & rsel_pmi & vbCr & "rsel_row - " & rsel_row

					If rsel_pmi <> "        " Then
						If InStr(PMIs_String, rsel_pmi) = 0 Then PMIs_String = PMIs_String & rsel_pmi & " "
					End If
				End If

				rsel_row = rsel_row + 1
				If rsel_row = 21 Then Exit Do
			Loop until rsel_pmi = "        "
		End If

	Next
	Call get_to_RKEY
	Call clear_line_of_text(4, 19)
	Call clear_line_of_text(5, 19)
	Call clear_line_of_text(6, 19)
	Call clear_line_of_text(6, 48)
	Call clear_line_of_text(6, 69)
	Call clear_line_of_text(7, 19)
	Call clear_line_of_text(7, 19)
	PMIs_String = trim(PMIs_String)
	MEMB_PMI_ARRAY = split(PMIs_String)
End Function

STATS_counter = UBound(MEMB_PMI_ARRAY)+1

const MMIS_case_numb_const 		= 0
const MMIS_pmi_const 			= 1
const DUPLICATE_pmi_const		= 2
const MMIS_ssn_const 			= 3
const MMIS_name_const			= 4
const MMIS_hc_prog_const		= 5
const MMIS_hc_elig_type_const	= 6
const MMIS_span_start_date_const= 7
const MMIS_span_status_const	= 8
const MMIS_span_end_date_const	= 9
const string_id_const			= 10
const MX_ref_numb_const 		= 11
const MMIS_last_const			= 12

Dim MMIS_HC_SPANS_ARRAY()
ReDim MMIS_HC_SPANS_ARRAY(MMIS_last_const, 0)

Call find_mmis_spans_past_4_months(MEMB_PMI_ARRAY, MMIS_case_numb_const, MMIS_pmi_const, DUPLICATE_pmi_const, MMIS_ssn_const, MMIS_name_const, MMIS_hc_prog_const, MMIS_hc_elig_type_const, MMIS_span_start_date_const, MMIS_span_status_const, MMIS_span_end_date_const, string_id_const, MX_ref_numb_const, MMIS_last_const, MMIS_HC_SPANS_ARRAY)

function find_mmis_spans_past_4_months(MEMB_PMI_ARRAY, MMIS_case_numb_const, MMIS_pmi_const, DUPLICATE_pmi_const, MMIS_ssn_const, MMIS_name_const, MMIS_hc_prog_const, MMIS_hc_elig_type_const, MMIS_span_start_date_const, MMIS_span_status_const, MMIS_span_end_date_const, string_id_const, MX_ref_numb_const, MMIS_last_const, MMIS_HC_SPANS_ARRAY)
	first_of_this_month = DatePart("m", date) & "/1/" & DatePart("yyyy", date)
	first_of_this_month = DateAdd("d", 0, first_of_this_month)
	four_months_ago = DateAdd("m", -4, first_of_this_month)

	span_count = 0
	For each MMIS_PMI in MEMB_PMI_ARRAY
		Call get_to_RKEY

		EMWriteScreen "I", 2, 19                                                    'read only
		EMWriteScreen MMIS_PMI, 4, 19                                             'enter through the PMI so it isn't case specific
		transmit

		rcip_duplicate_pmi = ""
		row = 1
		col = 1
		EMSearch "DUPLICATE PMI #", row, col
		If row <> 0 Then
			EMReadScreen rcip_duplicate_pmi, 8, row, col + 16
		End If

		Call write_value_and_transmit("RCIP", 1, 8)
		EMReadScreen rcip_ssn, 9, 5, 28
		rcip_ssn = left(rcip_ssn, 3) & "-" & mid(rcip_ssn, 4, 2) & "-" & right(rcip_ssn, 4)
		EMReadScreen rcip_name, 31, 3, 2
		EMReadScreen rcip_init, 1, 3, 33
		rcip_name = trim(rcip_name)
		Do
			rcip_name = replace(rcip_name, "  ", " ")
		Loop until InStr(rcip_name, "  ") = 0
		rcip_name = replace(rcip_name, " ", ", ")
		If rcip_init <> " " Then rcip_name = rcip_name & " " & rcip_init

		Call write_value_and_transmit("RELG", 1, 8)

		relg_row = 6                                'beginning of the list.
		Do

			EMReadScreen relg_prog, 2, relg_row, 10 'reading the prog and elig type information
			If relg_prog = "  " Then Exit Do
			If relg_prog = "* " Then Exit Do

			EMReadScreen relg_elig, 2, relg_row, 33
			EMReadScreen relg_case_num, 8, relg_row, 73 'reading the case number for this span
			EMReadScreen relg_status, 1, relg_row+1, 62
			EMReadScreen relg_start_dt, 8, relg_row+1, 15     'this is where the end date is
			EMReadScreen relg_end_dt, 8, relg_row+1, 36     'this is where the end date is

			record_elig = False
			If relg_status = "A" Then record_elig = True
			If relg_end_dt = "99/99/99" Then
				record_elig = True
			ElseIf DateDiff("d", four_months_ago, relg_end_dt) >= 0 Then
				record_elig = True
			End If

			If record_elig = True Then
				ReDim preserve MMIS_HC_SPANS_ARRAY(MMIS_last_const, span_count)
				MMIS_HC_SPANS_ARRAY(MMIS_pmi_const, span_count) = MMIS_PMI
				MMIS_HC_SPANS_ARRAY(MMIS_ssn_const, span_count) = rcip_ssn
				MMIS_HC_SPANS_ARRAY(MMIS_name_const, span_count) = rcip_name
				MMIS_HC_SPANS_ARRAY(DUPLICATE_pmi_const, span_count) = rcip_duplicate_pmi
				MMIS_HC_SPANS_ARRAY(MMIS_case_numb_const, span_count) = relg_case_num
				MMIS_HC_SPANS_ARRAY(MMIS_hc_prog_const, span_count) = relg_prog
				MMIS_HC_SPANS_ARRAY(MMIS_hc_elig_type_const, span_count) = relg_elig
				MMIS_HC_SPANS_ARRAY(MMIS_span_start_date_const, span_count) = relg_start_dt
				MMIS_HC_SPANS_ARRAY(MMIS_span_status_const, span_count) = relg_status
				MMIS_HC_SPANS_ARRAY(MMIS_span_end_date_const, span_count) = relg_end_dt
				MMIS_HC_SPANS_ARRAY(string_id_const, span_count) = MMIS_PMI & "-" & relg_case_num & "-" & relg_ssn & "-" & relg_prog & "-" & relg_elig' & "-" &  & "-" &  & "-" &
				span_count = span_count + 1
			End If

			relg_row = relg_row + 4         'next span on RELG'
			If relg_row = 22 Then           'this is the end of RELG and we need to go to a new page
				PF8
				relg_row = 6
				EMReadScreen end_of_list, 7, 24, 26     'This is the end of the list
				If end_of_list = "NO MORE" Then Exit Do
			End If
		Loop
	Next
End Function
Call navigate_to_MAXIS(MX_environment)                          'going back to MAXIS

Dialog1 = ""
If MMIS_HC_SPANS_ARRAY(MMIS_name_const, 0) <> "" Then
	dlg_len = 30 + 50*(UBound(MMIS_HC_SPANS_ARRAY,2)+1)
	BeginDialog Dialog1, 0, 0, 385, dlg_len, "MEMB " & MEMB_ref_numb & " - MMIS Span Information"
		Text 10, 10, 260, 10, "Member: MEMB " & MEMB_ref_numb & " - " & MEMB_last_name & ", " & MEMB_first_name & " - MAXIS PMI: " & MEMB_pmi
		ButtonGroup ButtonPressed
			OkButton 330, 5, 50, 15

		y_pos = 25
		For each_hc_span = 0 to UBound(MMIS_HC_SPANS_ARRAY, 2)
			GroupBox 10, y_pos, 375, 45, MMIS_HC_SPANS_ARRAY(MMIS_hc_prog_const, each_hc_span) & " - " & MMIS_HC_SPANS_ARRAY(MMIS_hc_elig_type_const, each_hc_span) & " on Case Number: " & MMIS_HC_SPANS_ARRAY(MMIS_case_numb_const, each_hc_span)
			If MMIS_HC_SPANS_ARRAY(DUPLICATE_pmi_const, each_hc_span) <> "" Then Text 200, y_pos, 150, 10, "DUPLICATE PMI: " & MMIS_HC_SPANS_ARRAY(DUPLICATE_pmi_const, each_hc_span)
			y_pos = y_pos + 15
			Text 25, y_pos, 85, 10, "Start Date: " & MMIS_HC_SPANS_ARRAY(MMIS_span_start_date_const, each_hc_span)
			Text 115, y_pos, 85, 10, "End Date: " & MMIS_HC_SPANS_ARRAY(MMIS_span_end_date_const, each_hc_span)
			Text 210, y_pos, 85, 10, "Status: " & MMIS_HC_SPANS_ARRAY(MMIS_span_status_const, each_hc_span)
			y_pos = y_pos + 15
			Text 25, y_pos, 70, 10, "PMI: " & MMIS_HC_SPANS_ARRAY(MMIS_pmi_const, each_hc_span)
			Text 115, y_pos, 80, 10, "SSN: " & MMIS_HC_SPANS_ARRAY(MMIS_ssn_const, each_hc_span)
			Text 210, y_pos, 140, 10, "Name: " & MMIS_HC_SPANS_ARRAY(MMIS_name_const, each_hc_span)
			y_pos = y_pos + 20
		Next
	EndDialog
Else
	dlg_len = 40 + 10*(UBound(MMIS_HC_SPANS_ARRAY,2)+1)
	BeginDialog Dialog1, 0, 0, 385, dlg_len, "MEMB " & MEMB_ref_numb & " - MMIS Span Information"
		Text 10, 10, 260, 10, "Member: MEMB " & MEMB_ref_numb & " - " & MEMB_last_name & ", " & MEMB_first_name & " - MAXIS PMI: " & MEMB_pmi
		Text 10, 25, 260, 10, "No HC Information found in MMIS. PMI(s) searched:"
		ButtonGroup ButtonPressed
			OkButton 330, 5, 50, 15
		y_pos = 35
		For each memb_pmi in MEMB_PMI_ARRAY
			Text 15, y_pos, 100, 10, " - " & memb_pmi
			y_pos = y_pos + 10
		Next
	EndDialog
End If

dialog Dialog1

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/13/2024
'--Tab orders reviewed & confirmed----------------------------------------------08/13/2024
'--Mandatory fields all present & Reviewed--------------------------------------08/13/2024
'--All variables in dialog match mandatory fields-------------------------------08/13/2024
'Review dialog names for content and content fit in dialog----------------------08/13/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------08/13/2024
'--Create a button to reference instructions------------------------------------08/13/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/13/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------08/13/2024
'--Out-of-County handling reviewed----------------------------------------------08/13/2024
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/13/2024
'--BULK - review output of statistics and run time/count (if applicable)--------08/13/2024
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------08/13/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/13/2024
'--Incrementors reviewed (if necessary)-----------------------------------------08/13/2024
'--Denomination reviewed -------------------------------------------------------08/13/2024
'--Script name reviewed---------------------------------------------------------08/13/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/13/2024

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------08/13/2024
'--comment Code-----------------------------------------------------------------08/13/2024
'--Update Changelog for release/update------------------------------------------08/13/2024
'--Remove testing message boxes-------------------------------------------------08/13/2024
'--Remove testing code/unnecessary code-----------------------------------------08/13/2024
'--Review/update SharePoint instructions----------------------------------------08/13/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
