'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - PREP DETAILS.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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



'DECLARATIONS ==============================================================================================================

Const memb_ref_numb_const 	= 0
Const memb_pmi_numb_const 	= 1
Const memb_ssn_const 		= 2
Const memb_age_const 		= 3
Const memb_name_const 		= 4
Const memb_active_hc_const	= 5
Const table_prog_1			= 6
Const table_type_1			= 7
Const table_prog_2			= 8
Const table_type_2			= 9
Const table_prog_3			= 10
Const table_type_3			= 11

Const unea_type_01_esists	= 20
Const unea_type_02_esists	= 21
Const unea_type_03_esists	= 22
Const unea_type_16_esists	= 23
Const unmatched_claim_numb	= 24
Const unea_VA_exists		= 25
Const unea_UC_exists		= 26

Const sves_qury_sent		= 40
Const second_qury_sent		= 41
Const sves_tpqy_response	= 42

Const memb_last_const 		= 70

Dim MEMBER_INFO_ARRAY()


Const va_case_numb_const 	= 0
Const va_ref_numb_const 	= 1
Const va_pers_name_const	= 2
Const va_pers_pmi_const		= 3
Const va_pers_ssn_const		= 4
Const va_inc_type_code_const 	= 5
Const va_inc_type_info_const	= 6
Const va_claim_numb_const 	= 7
Const va_prosp_inc_const 	= 8
Const va_last_const 		= 9

Dim VA_INCOME_ARRAY()
ReDim VA_INCOME_ARRAY(va_last_const, 0)

Const uc_case_numb_const 	= 0
Const uc_ref_numb_const 	= 1
Const uc_pers_name_const	= 2
Const uc_pers_pmi_const		= 3
Const uc_pers_ssn_const		= 4
Const uc_inc_type_code_const 	= 5
Const uc_inc_type_info_const	= 6
Const uc_claim_numb_const 	= 7
Const uc_prosp_inc_const 	= 8
Const uc_last_const 		= 9

Dim UC_INCOME_ARRAY()
ReDim UC_INCOME_ARRAY(va_last_const, 0)

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
Call MAXIS_case_number_finder(MAXIS_case_number)

ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

review_date = ep_revw_mo & "/1/" & ep_revw_yr
review_date = DateAdd("d", 0, review_date)

'Initial Dialog - Case number
Dialog1 = ""                                        'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 191, 135, "Application Received"
  EditBox 60, 35, 45, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    ' PushButton 90, 95, 95, 15, "Script Instructions", script_instructions_btn
    OkButton 80, 115, 50, 15
    CancelButton 135, 115, 50, 15
  Text 5, 10, 185, 20, "Review case details and display PREP findings"
  Text 5, 40, 50, 10, "Case Number:"
  Text 5, 55, 185, 10, "This case should be in PND2 status for this script to run."
  Text 5, 65, 185, 30, "If the programs requested on the application are not yet pending in MAXIS, cancel this script run, pend the case to PND2 status and run the script again."
EndDialog


'declare the SQL statement that will query the database
objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "' and [CaseNumber] = '" & MAXIS_case_number & "'"

'Creating objects for Access
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objRecordSet.Open objSQL, objConnection

all_hc_is_ABD = ""
SSA_income_exists = ""
JOBS_income_exists = ""
VA_income_exists = ""
BUSI_income_exists = ""
case_has_no_income = ""
case_has_EPD = ""

appears_ex_parte = True
all_hc_is_ABD = True
case_has_EPD = False
case_is_in_henn = False

ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)
memb_count = 0

objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

'Creating objects for Access
Set objELIGConnection = CreateObject("ADODB.Connection")
Set objELIGRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objELIGRecordSet.Open objELIGSQL, objELIGConnection

Do While NOT objELIGRecordSet.Eof

' If objELIGRecordSet("MajorProgram") = NULL
	memb_known = False
	For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
		If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then
			memb_known = True
			If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then
				MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= objELIGRecordSet("MajorProgram")
				MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= objELIGRecordSet("EligType")
			ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
				MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= objELIGRecordSet("MajorProgram")
				MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= objELIGRecordSet("EligType")
			ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
				MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= objELIGRecordSet("MajorProgram")
				MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= objELIGRecordSet("EligType")
			End If
			If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
			If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
			If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
			If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
			If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
			If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
			If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
			If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

			If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
			If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True
		End If
	Next

	If memb_known = False Then
		ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)

		MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
		MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
		MEMBER_INFO_ARRAY(memb_name_const, memb_count) 		= trim(objELIGRecordSet("Name"))
		MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
		MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))
		MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

		If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
		If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
		If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
		If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
		If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
		If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
		If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
		If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

		If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
		If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True

		memb_count = memb_count + 1
	End if
	objELIGRecordSet.MoveNext
Loop
objELIGRecordSet.Close
objELIGConnection.Close
Set objELIGRecordSet=nothing
Set objELIGConnection=nothing


SSA_income_exists = False
RR_income_exists = False
VA_income_exists = False
UC_income_exists = False
PRISM_income_exists = False
Other_UNEA_income_exists = False
JOBS_income_exists = False
BUSI_income_exists = False

objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

'Creating objects for Access
Set objIncomeConnection = CreateObject("ADODB.Connection")
Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

'This is the file path for the statistics Access database.
' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

Do While NOT objIncomeRecordSet.Eof
	If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then
		If objIncomeRecordSet("IncomeTypeCode") = "01" Then SSA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "02" Then SSA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "03" Then SSA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "16" Then SSA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "11" Then VA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "12" Then VA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "13" Then VA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "38" Then VA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True

		If objIncomeRecordSet("IncomeTypeCode") = "36" Then PRISM_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "37" Then PRISM_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "39" Then PRISM_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "40" Then PRISM_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "36" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "37" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "39" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "40" Then Other_UNEA_income_exists = True

		If objIncomeRecordSet("IncomeTypeCode") = "06" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "15" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "17" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "18" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "23" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "24" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "25" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "26" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "27" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "28" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "29" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "08" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "35" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "43" Then Other_UNEA_income_exists = True
		If objIncomeRecordSet("IncomeTypeCode") = "47" Then Other_UNEA_income_exists = True
	End If
	If objIncomeRecordSet("IncExpTypeCode") = "JOBS" Then JOBS_income_exists = True
	If objIncomeRecordSet("IncExpTypeCode") = "BUSI" Then BUSI_income_exists = True

	objIncomeRecordSet.MoveNext
Loop
objIncomeRecordSet.Close
objIncomeConnection.Close
Set objIncomeRecordSet=nothing
Set objIncomeConnection=nothing


If appears_ex_parte = True Then

	'check HC ER date in STAT/REVW
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
	If is_this_priv = True Then appears_ex_parte = False
	If is_this_priv = False Then
		Call write_value_and_transmit("X", 5, 71)
		EMReadScreen STAT_HC_ER_mo, 2, 8, 27
		EMReadScreen STAT_HC_ER_yr, 2, 8, 33
		If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False
	End If
End If

If appears_ex_parte = True Then
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	Call get_list_of_members

	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
	EMReadScreen case_pw, 7, 21, 14
	If left(case_pw, 4) = "X127" Then case_is_in_henn = True
	If case_is_in_henn = False then  appears_ex_parte = False
	If case_active = False Then appears_ex_parte = False
	If ma_status <> "ACTIVE" and msp_status <> "ACTIVE" Then appears_ex_parte = False
	' If msp_status <> "ACTIVE" Then appears_ex_parte = False

	' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then
End If

case_has_no_income = False
If SSA_income_exists = False and RR_income_exists = False and VA_income_exists = False and UC_income_exists = False and PRISM_income_exists = False and Other_UNEA_income_exists = False and JOBS_income_exists = False and BUSI_income_exists = False Then case_has_no_income = True

If appears_ex_parte = True Then
	If Other_UNEA_income_exists = True OR JOBS_income_exists = True OR BUSI_income_exists = True Then
		If mfip_status = "ACTIVE" Then
			' If all_hc_is_ABD = True and case_has_EPD = False Then
			appears_ex_parte = True
		ElseIf snap_status = "ACTIVE" Then
			'find income
			appears_ex_parte = True
		Else
			' If Other_UNEA_income_exists = True Then appears_ex_parte = False
			' If JOBS_income_exists = True Then appears_ex_parte = False
			' If BUSI_income_exists = True Then appears_ex_parte = False
			appears_ex_parte = False
		End If
	End If
End If

If appears_ex_parte = True Then
	'For each case that is indicated as potentially ExParte, we are going to take preperation actions
	last_va_count = va_count
	last_uc_count = uc_count

	'Find if there is a claim number that is not associated with the persons SSN
	'Read if VA income is on UNEA to add that person to a list to verify VA
	'Read if UC income is on UNEA to add that person to a list to verify UC
	Call find_unea_information

End If

Call back_to_SELF
' objRecordSet.Update "SelectExParte", appears_ex_parte
prep_status = date
If appears_ex_parte = False Then
	prep_status = "Not Ex Parte"
	' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then prep_status = "SNAP/MFIP"
End If

