'STATS GATHERING=============================================================================================================
name_of_script = "MISC - VERIFY EX PARTE.vbs"       'TO DO - UPDATE SCRIPT TITLE. REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
'TO DO - update changelog if needed
CALL changelog_update("05/25/2023", "Initial version.", "Mark Riegel, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabs the MAXIS case number automatically

'Gather Case Number and the form processed
Dialog1 = ""

BeginDialog Dialog1, 0, 0, 366, 300, "Health Care Evaluation"
  EditBox 80, 200, 50, 15, MAXIS_case_number
  DropListBox 80, 220, 275, 45, "Select One..."+chr(9)+"No Form - Ex Parte Determination", HC_form_name
  DropListBox 265, 240, 90, 45, "No"+chr(9)+"Yes"+chr(9)+"N/A - Ex Parte Process", ltc_waiver_request_yn
  EditBox 80, 260, 50, 15, form_date
  EditBox 80, 280, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 280, 50, 15
    CancelButton 305, 280, 50, 15
    PushButton 295, 35, 50, 15, "Instructions", instructions_btn
    PushButton 295, 50, 50, 15, "Video Demo", video_demo_btn
  Text 105, 10, 120, 10, "Health Care Evaluation Script"
  Text 20, 40, 255, 20, "This script is to be run once MAXIS STAT panels have been updated with all accurate information from a Health Care Application Form."
  Text 20, 65, 255, 25, "If information displayed in this script is inaccurate, this means the information entered into STAT requires update. Cancel the script run and update STAT panels before running the script again."
  Text 20, 95, 255, 10, "The information and coding in STAT will directly pull into the script details:"
  Text 35, 105, 250, 10, "- Panels coded as needing verification will show up as verifications needed."
  Text 35, 115, 250, 10, "- Income amounts will be pulled from JOBS / UNEA / BUSI / ect panels"
  Text 40, 125, 150, 10, "and cannot be updated in the script dialogs."
  Text 35, 135, 250, 10, "- Asset amounts will be pulled from ACCT / CASH / SECU / ect panels and"
  Text 40, 145, 175, 10, "cannot be updated in the script dialogs."
  Text 35, 155, 250, 10, "- The details in STAT Panels should be accurate and the script serves as a"
  Text 40, 165, 245, 10, "secondary review of information that makes an eligibility determinations."
  Text 15, 180, 300, 10, "IF THE CASE INFORMATION IS WRONG IN THE SCRIPT, IT IS WRONG IN THE SYSTEM"
  Text 25, 205, 50, 10, "Case Number:"
  Text 15, 225, 60, 10, "Health Care Form:"
  Text 80, 245, 185, 10, "Does this form qualify to request LTC/Waiver Services?"
  Text 35, 265, 40, 10, "Form Date:"
  Text 15, 285, 60, 10, "Worker Signature:"
  GroupBox 10, 25, 345, 170, "Health Care Processing"
  Text 135, 265, 110, 10, "For ex parte, use processing date"
EndDialog

DO
	DO
	   	err_msg = ""
	   	Dialog Dialog1
	   	cancel_without_confirmation

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

If HC_form_name = "No Form - Ex Parte Determination" Then
	SQL_Case_Number = right("00000000" & MAXIS_case_number, 8)
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
	If ex_parte_determination_from_SQL = 0 Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") was previously determined to not meet Ex Parte criteria.")
	review_month_from_SQL = objRecordSet("HCEligReviewDate")
	If review_month_from_SQL = "" Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") is not listed on the Ex Parte Data Table and cannot be processed as Ex Parte.")
	review_month_from_SQL = DateAdd("d", 0, review_month_from_SQL)
	Call convert_date_into_MAXIS_footer_month(review_month_from_SQL, er_month, er_year)
	If DateDiff("d", date, review_month_from_SQL) =<0 Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") has a HC ER listed in the Ex Parte Data Table as " & er_month & "/" & er_year & ", which is in the past and cannot be processed as Ex Parte.")
	If DateDiff("d", date, review_month_from_SQL) > 90 Then call script_end_procedure_with_error_report("This case (" & MAXIS_case_number & ") has a HC ER listed in the Ex Parte Data Table as " & er_month & "/" & er_year & ", which is too far in the future to be processed as Ex Parte.")
	ex_parte_phase = ""
	If DateDiff("d", date, review_month_from_SQL) < 32 Then
		ex_parte_phase = "Phase 2"
	Else
		ex_parte_phase = "Phase 1"
	End If
	review_month_01 = er_month & "/" & er_year
	review_month_02 = er_month & "/" & er_year

	objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & SQL_Case_Number & "'"

	'Creating objects for Access
	Set objELIGConnection = CreateObject("ADODB.Connection")
	Set objELIGRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objELIGRecordSet.Open objELIGSQL, objELIGConnection

	Do While NOT objELIGRecordSet.Eof

		If name_01 = "" Then
			name_01 = trim(objELIGRecordSet("Name"))
			PMI_01 = trim(objELIGRecordSet("PMINumber"))

			If objELIGRecordSet("MajorProgram") = "MA" Then
				MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
			Else
				MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
				MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
			End If
		ElseIf PMI_01 = trim(objELIGRecordSet("PMINumber")) Then
			If objELIGRecordSet("MajorProgram") = "MA" Then
				MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
			Else
				MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
				MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
			End If
		ElseIf name_02 = "" Then
			name_02 = trim(objELIGRecordSet("Name"))
			PMI_02 = trim(objELIGRecordSet("PMINumber"))

			If objELIGRecordSet("MajorProgram") = "MA" Then
				MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
			Else
				MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
				MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
			End If
		ElseIf PMI_02 = trim(objELIGRecordSet("PMINumber")) Then
			If objELIGRecordSet("MajorProgram") = "MA" Then
				MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
			Else
				MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
				MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
			End If
		End If
		objELIGRecordSet.MoveNext
	Loop

	objELIGRecordSet.Close
	objELIGConnection.Close
	Set objELIGRecordSet=nothing
	Set objELIGConnection=nothing

	objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & SQL_Case_Number & "'"

	'Creating objects for Access
	Set objIncomeConnection = CreateObject("ADODB.Connection")
	Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

	Do While NOT objIncomeRecordSet.Eof
		panel_name_from_SQL = objIncomeRecordSet("IncExpTypeCode")
		income_type_code_from_SQL = objIncomeRecordSet("IncomeTypeCode")
		prosp_amt_from_SQL = objIncomeRecordSet("ProspAmount")
        income_description_from_SQL = objIncomeRecordSet("Description")

		If trim(objIncomeRecordSet("PersonID")) = PMI_01 Then
			sent_date_01 = objIncomeRecordSet("QURY_Sent")
			' If income_type_code_from_SQL = "01" or income_type_code_from_SQL = "02" or income_type_code_from_SQL = "03" Then
			If income_type_code_from_SQL = "01" Then bndx_amount_01 = income_description_from_SQL & " - " & prosp_amt_from_SQL
			If income_type_code_from_SQL = "02" Then bndx_amount_01 = income_description_from_SQL & " - " & prosp_amt_from_SQL
			If income_type_code_from_SQL = "03" Then sdxs_amount_01 = income_description_from_SQL & " - " & prosp_amt_from_SQL
            If panel_name_from_SQL = "JOBS" Then JOBS_01 = "Yes"
			If income_type_code_from_SQL <> "01" AND income_type_code_from_SQL <> "02" AND income_type_code_from_SQL <> "03" Then other_UNEA_types_01 = other_UNEA_types_01 & " " & income_type_code_from_SQL
		End If

        If trim(objIncomeRecordSet("PersonID")) = PMI_02 Then
			sent_date_02 = objIncomeRecordSet("QURY_Sent")
			' If income_type_code_from_SQL = "01" or income_type_code_from_SQL = "02" or income_type_code_from_SQL = "03" Then
			If income_type_code_from_SQL = "01" Then bndx_amount_02 = income_description_from_SQL & " - " & prosp_amt_from_SQL
			If income_type_code_from_SQL = "02" Then bndx_amount_02 = income_description_from_SQL & " - " & prosp_amt_from_SQL
			If income_type_code_from_SQL = "03" Then sdxs_amount_02 = income_description_from_SQL & " - " & prosp_amt_from_SQL
            If panel_name_from_SQL = "JOBS" Then JOBS_02 = "Yes"
			If income_type_code_from_SQL <> "01" AND income_type_code_from_SQL <> "02" AND income_type_code_from_SQL <> "03" Then other_UNEA_types_02 = other_UNEA_types_02 & " " & income_type_code_from_SQL
		End If

		objIncomeRecordSet.MoveNext
	Loop

	objIncomeRecordSet.Close
	objIncomeConnection.Close
	Set objIncomeRecordSet=nothing
	Set objIncomeConnection=nothing

	Dialog1 = ""

    'TO DO - add functionality to determine ex parte phase based on case number
    BeginDialog Dialog1, 0, 0, 556, 385, "Phase 1 - Ex Parte Determination"
        GroupBox 10, 310, 455, 50, "Ex Parte Determination"
        Text 20, 325, 80, 10, "Ex Parte Determination:"
        DropListBox 100, 320, 115, 50, ""+chr(9)+"Appears Ex Parte"+chr(9)+"Cannot be Processed as Ex Parte", ex_parte_determination
        Text 20, 340, 205, 10, "If case cannot be processed as ex parte, provide explanation:"
        EditBox 220, 335, 235, 15, ex_parte_denial_explanation
        Text 15, 365, 70, 10, "Worker Signature:"
        EditBox 80, 360, 110, 15, worker_signature
        GroupBox 10, 0, 455, 25, "Case Information"
            Text 15, 10, 50, 10, "Case Number:"
            Text 65, 10, 70, 10, MAXIS_case_number
            Text 140, 10, 50, 10, "Review Month:"
            Text 190, 10, 65, 10, review_month_01
        ButtonGroup ButtonPressed
            OkButton 440, 365, 50, 15
            CancelButton 500, 365, 50, 15
        GroupBox 10, 25, 220, 35, "Person 1 - Information"
            Text 15, 35, 25, 10, "Name:"
            Text 40, 35, 180, 10, name_01
            Text 15, 45, 20, 10, "PMI:"
            Text 35, 45, 70, 10, PMI_01
        GroupBox 10, 60, 220, 45, "Person 1 - TPQY Information"
            'TO DO - determine functionality to add multiple claim numbers
            ' Text 15, 75, 50, 10, "Claim Number:"
            ' Text 65, 75, 50, 10, claim_number_01
            Text 15, 70, 35, 10, "Sent Date:"
            Text 55, 70, 50, 10, sent_date_01
            Text 120, 70, 45, 10, "Return Date:"
            Text 165, 70, 55, 10, return_date_01
            Text 15, 90, 50, 10, "SDXS Amount:"
            Text 65, 90, 155, 10, sdxs_amount_01
            Text 15, 80, 50, 10, "BNDX Amount:"
            Text 65, 80, 155, 10, bndx_amount_01
        ButtonGroup ButtonPressed
            Text 480, 5, 70, 10, "--- INSTRUCTIONS ---"
            PushButton 490, 15, 55, 15, "instructions", instructions_button
            Text 495, 40, 45, 10, "--- POLICY ---"
            PushButton 490, 50, 55, 15, policy_1, policy_1_button
            PushButton 490, 65, 55, 15, policy_2, policy_2_button
            PushButton 490, 80, 55, 15, policy_3, policy_3_button
            Text 490, 105, 55, 10, "--- NAVIGATE ---"
            PushButton 490, 115, 25, 15, "ACCI", acci_button
            PushButton 515, 115, 25, 15, "BILS", bils_button
            PushButton 490, 130, 25, 15, "BUDG", budg_button
            PushButton 515, 130, 25, 15, "BUSI", busi_button
            PushButton 490, 145, 25, 15, "DISA", disa_button
            PushButton 515, 145, 25, 15, "EMMA", emma_button
            PushButton 490, 160, 25, 15, "FACI", faci_button
            PushButton 515, 160, 25, 15, "HCMI", hcmi_button
            PushButton 490, 175, 25, 15, "IMIG", imig_button
            PushButton 515, 175, 25, 15, "INSA", insa_button
            PushButton 490, 190, 25, 15, "JOBS", jobs_button
            PushButton 515, 190, 25, 15, "LUMP", lump_button
            PushButton 490, 205, 25, 15, "MEDI", medi_button
            PushButton 515, 205, 25, 15, "MEMB", memb_button
            PushButton 490, 220, 25, 15, "MEMI", memi_button
            PushButton 515, 220, 25, 15, "PBEN", pben_button
            PushButton 490, 235, 25, 15, "PDED", pded_button
            PushButton 515, 235, 25, 15, "REVW", revw_button
            PushButton 490, 250, 25, 15, "SPON", spon_button
            PushButton 515, 250, 25, 15, "STWK", stwk_button
            PushButton 490, 265, 25, 15, "UNEA", unea_button
        GroupBox 10, 105, 220, 200, "Person 1 - Add'l Information"
            Text 15, 115, 35, 10, "MEDI Info: "
            Text 85, 115, 135, 10, medi_info_01
            Text 15, 130, 70, 10, "MEDI - Part A Exists:"
            Text 85, 130, 135, 10, MEDI_part_a_01
            Text 15, 145, 70, 10, "MEDI - Part B Exists:"
            Text 85, 145, 135, 10, MEDI_part_b_01
            Text 15, 175, 65, 10, "Other UNEA Types:"
            Text 85, 175, 135, 10, other_UNEA_types_01
            Text 15, 160, 50, 10, "JOBS Exists:"
            Text 85, 160, 135, 10, JOBS_01
            Text 15, 190, 60, 10, "MAXIS MA Basis:"
            Text 85, 190, 135, 10, MAXIS_MA_basis_01
            Text 15, 205, 60, 10, "MAXIS MSP Prog:"
            Text 85, 205, 135, 10, MAXIS_msp_prog_01
            Text 15, 220, 65, 10, "MAXIS MSP Basis:"
            Text 85, 220, 135, 10, MAXIS_msp_basis_01
            Text 15, 235, 55, 10, "MMIS MA Basis:"
            Text 85, 235, 135, 10, MMIS_ma_basis_01
            Text 15, 250, 60, 10, "MMIS MSP Prog:"
            Text 85, 250, 135, 10, MMIS_msp_prog_01
            Text 15, 265, 60, 10, "MMIS MSP Basis:"
            Text 85, 265, 135, 10, MMIS_msp_basis_01
		If name_02 <> "" Then
			GroupBox 245, 25, 220, 35, "Person 2 - Information"
                Text 250, 35, 25, 10, "Name:"
                Text 275, 35, 185, 10, name_02
                Text 250, 45, 20, 10, "PMI:"
                Text 275, 45, 60, 10, PMI_02
			GroupBox 245, 60, 220, 45, "Person 2 - TPQY Information"
                'TO DO - determine functionality to add multiple claim numbers
				' Text 250, 75, 50, 10, "Claim Number:"
				' Text 300, 75, 50, 10, claim_number_02
                Text 250, 70, 35, 10, "Sent Date:"
                Text 290, 70, 50, 10, sent_date_02
                Text 355, 70, 45, 10, "Return Date:"
                Text 400, 70, 55, 10, return_date_02
                Text 250, 90, 50, 10, "SDXS Amount:"
                Text 300, 90, 155, 10, sdxs_amount_02
                Text 250, 80, 50, 10, "BNDX Amount:"
                Text 300, 80, 155, 10, bndx_amount_02
            GroupBox 245, 105, 220, 200, "Person 2 - Add'l Information"
                Text 255, 115, 35, 10, "MEDI Info: "
                Text 325, 115, 135, 10, medi_info_02
                Text 255, 130, 70, 10, "MEDI - Part A Exists:"
                Text 325, 130, 135, 10, MEDI_part_a_02
                Text 255, 145, 70, 10, "MEDI - Part B Exists:"
                Text 325, 145, 135, 10, MEDI_part_b_02
                Text 255, 175, 65, 10, "Other UNEA Types:"
                Text 325, 175, 135, 10, other_UNEA_types_02
                Text 255, 160, 50, 10, "JOBS Exists:"
                Text 325, 160, 135, 10, JOBS_02
                Text 255, 190, 60, 10, "MAXIS MA Basis:"
                Text 325, 190, 135, 10, MAXIS_MA_basis_02
                Text 255, 205, 60, 10, "MAXIS MSP Prog:"
                Text 325, 205, 135, 10, MAXIS_msp_prog_02
                Text 255, 220, 65, 10, "MAXIS MSP Basis:"
                Text 325, 220, 135, 10, MAXIS_msp_basis_02
                Text 255, 235, 55, 10, "MMIS MA Basis:"
                Text 325, 235, 135, 10, MMIS_ma_basis_02
                Text 255, 250, 60, 10, "MMIS MSP Prog:"
                Text 325, 250, 135, 10, MMIS_msp_prog_02
                Text 255, 265, 60, 10, "MMIS MSP Basis:"
                Text 325, 265, 135, 10, MMIS_msp_basis_02
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

            'Add placeholder link to script instructions - To DO - update with correct link
            If ButtonPressed = instructions_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"

            'Add placeholder links for policy buttons - TO DO - update with correct links
            If ButtonPressed = policy_1_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
            If ButtonPressed = policy_2_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"
            If ButtonPressed = policy_3_button Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/human-services"


            'Add validation to ensure ex parte determination is made
            If ex_parte_determination = "" THEN err_msg = err_msg & vbCr & "* You must make an ex parte determination."

            'Add validation that if ex parte approved, then explanation should be blank
            If ex_parte_determination = "Appears Ex Parte" AND trim(ex_parte_denial_explanation) <> "" THEN err_msg = err_msg & vbCr & "* The explanation for denial field should be blank since ex parte has been approved."

            'Add validation that if ex parte denied, then explanation must be provided
            If ex_parte_determination = "Cannot be Processed as Ex Parte" AND trim(ex_parte_denial_explanation) = "" THEN err_msg = err_msg & vbCr & "* You must provide an explanation for the ex parte denial."

            'Add validation to ensure worker signature is not blank
            IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please include your worker signature."

            'Error message handling
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = "" AND ButtonPressed = -1
        'Add to all dialogs where you need to work within BLUEZONE
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
    'End dialog section-----------------------------------------------------------------------------------------------
End If

'Checks to see if in MAXIS
Call check_for_MAXIS(False)

'Ensure starting at SELF so that writing to CASE NOTE works properly
CALL back_to_SELF()

'Navigate to STAT, REVW, and open HC Renewal Window with instructions
CALL navigate_to_MAXIS_screen("STAT", "REVW")
CALL write_value_and_transmit("X", 5, 71)

'Read data from HC renewal screen to determine what changes the worker needs to complete and then use to validate changes
'TO DO - update variables to match/pull from SQL data table. This data should be used as baseline/reference point for validation.
EMReadScreen income_renewal_date, 8, 7, 27
EMReadScreen elig_renewal_date, 8, 8, 27
EMReadScreen HC_ex_parte_determination, 1, 9, 27
EMReadScreen income_asset_renewal_date, 8, 7, 71
EMReadScreen exempt_6_mo_ir_form, 1, 8, 71
EMReadScreen ex_parte_renewal_month_year, 7, 9, 71

'Convert variables from HC Renewal Screen to dates
'TO DO - update once we have variables from SQL table as this may no longer be needed
income_renewal_date = replace(income_renewal_date, " ", "/")
elig_renewal_date = replace(elig_renewal_date, " ", "/")
income_asset_renewal_date = replace(income_asset_renewal_date, " ", "/")
ex_parte_renewal_month_year = replace(ex_parte_renewal_month_year, " ", "/")

'Dim variables for use in function so that it updates variables on first run of do loop
dim check_income_renewal_date
dim check_elig_renewal_date
dim check_HC_ex_parte_determination
dim check_income_asset_renewal_date
dim check_exempt_6_mo_ir_form
dim check_ex_parte_renewal_month_year


'Create function to read HC Renewal Screen to verify changes made by worker. Also converts variables to dates where applicable.
'TO DO - figure out why this does not execute immediately when called in do loop
Function check_hc_renewal_updates()
    EMReadScreen check_income_renewal_date, 8, 7, 27
    check_income_renewal_date = replace(check_income_renewal_date, " ", "/")
    EMReadScreen check_elig_renewal_date, 8, 8, 27
    check_elig_renewal_date = replace(check_elig_renewal_date, " ", "/")
    EMReadScreen check_HC_ex_parte_determination, 1, 9, 27
    EMReadScreen check_income_asset_renewal_date, 8, 7, 71
    check_income_asset_renewal_date = replace(check_income_asset_renewal_date, " ", "/")
    EMReadScreen check_exempt_6_mo_ir_form, 1, 8, 71
    EMReadScreen check_ex_parte_renewal_month_year, 7, 9, 71
    check_ex_parte_renewal_month_year = replace(check_ex_parte_renewal_month_year, " ", "/")
End Function

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
            If ButtonPressed = hc_renewal_button Then
                EMReadScreen income_renewal_date, 8, 7, 27
                EMReadScreen elig_renewal_date, 8, 8, 27
                EMReadScreen HC_ex_parte_determination, 1, 9, 27
                EMReadScreen income_asset_renewal_date, 8, 7, 71
                EMReadScreen exempt_6_mo_ir_form, 1, 8, 71
                EMReadScreen ex_parte_renewal_month_year, 7, 9, 71
            End If

            'Validate Elig Renewal Date to ensure it is set for 1 year from current Elig Renewal Date
            If check_elig_renewal_date <> DateAdd("Y", 1, elig_renewal_date) THEN err_msg = err_msg & vbCr & "* The Elig Renewal Date should be set for 1 year from the current renewal month and year."

            'Validate Income/Asset Renewal Date to ensure it is the same as the Elig Renewal Date or set for 6 months from original Elig Renewal Date for cases with a spenddown:
            'TO DO - determine how to determine if meets spenddown?
            If check_income_asset_renewal_date <> DateAdd("Y", 1, income_asset_renewal_date) OR check_income_asset_renewal_date <> DateAdd("M", 6, income_asset_renewal_date) THEN err_msg = err_msg & vbCr & "* The Income/Asset Renewal Date should be be the same as the Elig Renewal Date. For cases with a spenddown that do not meet an exception listed in EPM 2.3.4.2 MA-ABD Renewals, enter a date six months from the original ELIG Renewal Date."

            'Validate that Exempt from 6 Mo IR is set to N
            If check_exempt_6_mo_ir_form <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for Exempt from 6 Mo IR."

            'Validate that ExParte field updated to Y
            If check_HC_ex_parte_determination <> "Y" THEN err_msg = err_msg & vbCr & "* You must enter 'Y' for ExParte."

            'Validate that ExParte Renewal Month is correct
            'TO DO - add validation to ensure that date updated in HC renewal screen is the same as date provided in SQL table
            If check_ex_parte_renewal_month_year = "__ ____" THEN err_msg = err_msg & vbCr & "* You must enter the month and year for the Ex Parte renewal month."

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
            If ButtonPressed = hc_renewal_button Then
                EMReadScreen income_renewal_date, 8, 7, 27
                EMReadScreen elig_renewal_date, 8, 8, 27
                EMReadScreen HC_ex_parte_determination, 1, 9, 27
                EMReadScreen income_asset_renewal_date, 8, 7, 71
                EMReadScreen exempt_6_mo_ir_form, 1, 8, 71
                EMReadScreen ex_parte_renewal_month_year, 7, 9, 71
            End If

            'Validation to ensure that elig renewal date has not changed
            If check_elig_renewal_date <> elig_renewal_date THEN err_msg = err_msg & vbCr & "* The Elig Renewal Date should not have been changed. It should remain " & elig_renewal_date & "."

            'Validation for Income/Asset Renewal Date to ensure that information has not changed
            If check_income_asset_renewal_date <> income_asset_renewal_date THEN err_msg = err_msg & vbCr & "* The Income/Asset Renewal Date should not have been changed. It should remain " & income_asset_renewal_date & "."

            'Validation to ensure that Exempt from 6 Mo IR is set to N
            If check_exempt_6_mo_ir_form <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for Exempt from 6 Mo IR."

            'Validation to ensure that ExParte field updated to N
            If check_HC_ex_parte_determination <> "N" THEN err_msg = err_msg & vbCr & "* You must enter 'N' for ExParte."

            'Validate that ExParte Renewal Month is correct
            'TO DO - add validation to ensure that date updated in HC renewal screen is the same as date provided in SQL table
            If check_ex_parte_renewal_month = "__ ____" THEN err_msg = err_msg & vbCr & "* You must enter the month and year for the Ex Parte renewal month."

            'Error message handling
            IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
            'Add to all dialogs where you need to work within BLUEZONE
            CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
End If


'If ex parte approved, create TIKL for 1st of processing month which is renewal month - 1
'TO DO - confirm TIKL information is correct
If ex_parte_determination = "Appears Ex Parte" Then Call create_TIKL("Phase 1 - The case has been evaluated for ex parte and has been approved based on the information provided.", 10, DateAdd("M", -1, ex_parte_renewal_month_year), False, TIKL_note_text)

'Navigate to and start a new CASE NOTE
Call start_a_blank_case_note

'Add title to CASE NOTE
CALL write_variable_in_case_note("*** EX PARTE DETERMINATION - " & UCASE(ex_parte_determination) & " ***")

'For ex parte approval, write information to case note
If ex_parte_determination = "Appears Ex Parte" Then
    CALL write_variable_in_case_note("Phase 1 - The case has been evaluated for ex parte and has been approved based on the information provided. The case meets one of the criteria below.")
    CALL write_variable_in_case_note("An MA-ABD enrollees will be ex parte renewed if their only source of income is:")
    CALL write_bullet_and_variable_in_case_note("* ", "Supplemental Security Income (SSI), even if the benefit amount is zero")
    CALL write_bullet_and_variable_in_case_note("* ", "Retirement, Survivors, and Disability Insurance (RSDI)")
    CALL write_bullet_and_variable_in_case_note("* ", "SSI + RSDI")
    CALL write_bullet_and_variable_in_case_note("* ", "Railroad Retirement Benefits (RRB)")
    CALL write_bullet_and_variable_in_case_note("* ", "RSDI + RRB")
    'TO DO - add additional language listing what would qualify for ex parte?
    'TO DO - add additional case details - case number, renewal info, etc?
End If


'For ex parte denial, write information to case note
If ex_parte_determination = "Cannot be Processed as Ex Parte" Then
    CALL write_variable_in_case_note("Phase 1 - The case has been evaluated for ex parte and has been denied based on the information provided.")
    CALL write_bullet_and_variable_in_case_note("Reason for Denial:", ex_parte_denial_explanation)
    'TO DO - add additional language listing what would qualify for ex parte?
    'TO DO - add additional case details - case number, renewal info, etc?
End If

'Add worker signature
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)


'Script end procedure
script_end_procedure("Success! The ex parte determination has been added to the CASE NOTE")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/23/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/23/2023
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------


