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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARY BLOCK================================================================================================


const case_numb_const = 0
const worker_numb_const = 1
const prog_info_const = 2
const pmi_numb_const = 3
const ref_numb_const = 4
const age_const = 5
const marital_const = 6
const MFIP_memb_code_const = 7
const rsdi_1_amt_const = 8
const rsdi_2_amt_const = 9
const ssi_amt_const = 10
const last_const = 11

Dim CASE_NUMBER_ARRAY()
ReDim CASE_NUMBER_ARRAY(2, 0)

Dim PERSON_ARRAY()
ReDim PERSON_ARRAY(last_const, 0)

Call check_for_MAXIS(true)

MAXIS_footer_month = "09"
MAXIS_footer_year = "24"

'file_url = "C:\Users\calo001\OneDrive - Hennepin County/Projects/2024-10 Legislative Changes/2024-08-08 MFIP Active Assessment.xlsx"
file_url = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Eligibility Summary\Lists of Cases\2024-09-24 MFIP Approvals.xlsx"
visible_status = True
alerts_status = True
Call excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)

'ObjExcel.worksheets("RSDI Impacted Cases").Activate
excel_row = 2
Do
	MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

	sig_change = False
	call navigate_to_MAXIS_screen("ELIG", "MFIP")
	EMReadScreen sig_change_check, 4, 3, 38				'looking to see if the significant change panel is on this case
	If sig_change_check = "MFSC" Then
		'this is important because the command line is in a different place on the sig change panel so this call is slightly different
		Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
	Else
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
	End If

	EMReadScreen significant_change_check, 18, 4, 3
	If significant_change_check = "SIGNIFICANT CHANGE" Then sig_change = True

	ObjExcel.Cells(excel_row, 2).Value = sig_change
	Call back_to_SELF
	excel_row = excel_row + 1

	' prog_info = trim(ObjExcel.Cells(excel_row, 2).Value)
	' Call back_to_SELF

	' If InStr(prog_info, "MF A") <> 0 Then
	' 	call navigate_to_MAXIS_screen("ELIG", "MFIP")
	' 	EMReadScreen sig_change_check, 4, 3, 38				'looking to see if the significant change panel is on this case
	' 	If sig_change_check = "MFSC" Then
	' 		'this is important because the command line is in a different place on the sig change panel so this call is slightly different
	' 		Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
	' 	Else
	' 		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
	' 	End If
	' 	Call write_value_and_transmit("MFSM", 20, 71)
	' 	EMReadScreen hrf_reporting, 10, 8, 31
	' 	objExcel.cells(excel_row, 3).value = trim(hrf_reporting)
	' ElseIf InStr(prog_info, "DW A") <> 0 Then
	' 	objExcel.cells(excel_row, 3).value = "DWP"
	' Else
	' 	objExcel.cells(excel_row, 3).value = "PENDING"
	' End If

	' excel_row = excel_row + 1
Loop Until trim(ObjExcel.Cells(excel_row, 1).Value) = ""
Call script_end_procedure("Done with reporting status")

'Read all of the case numbers and worker numbers
ObjExcel.worksheets("MFIP Only Cases").Activate
excel_row = 3716			'CASE READING START
case_count = 0
Do
	ReDim preserve CASE_NUMBER_ARRAY(2, case_count)
	CASE_NUMBER_ARRAY(worker_numb_const, case_count) = trim(ObjExcel.Cells(excel_row, 1).Value)
	CASE_NUMBER_ARRAY(case_numb_const, case_count) = trim(ObjExcel.Cells(excel_row, 2).Value)
	CASE_NUMBER_ARRAY(prog_info_const, case_count) = trim(ObjExcel.Cells(excel_row, 5).Value)
	case_count = case_count + 1
	excel_row = excel_row + 1
Loop Until trim(ObjExcel.Cells(excel_row, 1).Value) = ""

pers_count = 0
excel_row = 12750		'PERSON RECORDING START
ObjExcel.worksheets("Income Assessment").Activate
For each_case = 0 to Ubound(CASE_NUMBER_ARRAY, 2)
	Call back_to_SELF
	MAXIS_case_number = CASE_NUMBER_ARRAY(case_numb_const, each_case)
	first_pers_index = pers_count

	If InStr(CASE_NUMBER_ARRAY(prog_info_const, each_case), "MF A") <> 0 Then
		call navigate_to_MAXIS_screen("ELIG", "MFIP")
		EMReadScreen sig_change_check, 4, 3, 38				'looking to see if the significant change panel is on this case
		If sig_change_check = "MFSC" Then
			'this is important because the command line is in a different place on the sig change panel so this call is slightly different
			Call find_last_approved_ELIG_version(19, 78, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		Else
			Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)
		End If

		' MsgBox "MFIP ELIG?"
		row = 7
		Do
			EMReadScreen ref_numb, 2, row, 6
			EMReadScreen member_code, 1, row, 36

			ReDim preserve PERSON_ARRAY(last_const, pers_count)
			PERSON_ARRAY(case_numb_const, pers_count) = CASE_NUMBER_ARRAY(case_numb_const, each_case)
			PERSON_ARRAY(worker_numb_const, pers_count) = CASE_NUMBER_ARRAY(worker_numb_const, each_case)
			PERSON_ARRAY(prog_info_const, pers_count) = CASE_NUMBER_ARRAY(prog_info_const, each_case)
			PERSON_ARRAY(ref_numb_const, pers_count) = ref_numb
			PERSON_ARRAY(MFIP_memb_code_const, pers_count) = member_code

			row = row + 1
			pers_count = pers_count + 1
			EMReadScreen next_ref_numb, 2, row, 6
			If row = 18 then
				PF8
				row = 7
			End if
		Loop until next_ref_numb = "  "
	ElseIf InStr(CASE_NUMBER_ARRAY(prog_info_const, each_case), "DW A") <> 0 Then
		call navigate_to_MAXIS_screen("ELIG", "DWP ")
		Call find_last_approved_ELIG_version(20, 79, elig_version_number, elig_version_date, elig_version_result, approved_version_found)

		' MsgBox "DWP ELIG?"
		row = 7
		Do
			EMReadScreen ref_numb, 2, row, 5
			EMReadScreen member_code, 1, row, 35

			ReDim preserve PERSON_ARRAY(last_const, pers_count)
			PERSON_ARRAY(case_numb_const, pers_count) = CASE_NUMBER_ARRAY(case_numb_const, each_case)
			PERSON_ARRAY(worker_numb_const, pers_count) = CASE_NUMBER_ARRAY(worker_numb_const, each_case)
			PERSON_ARRAY(prog_info_const, pers_count) = CASE_NUMBER_ARRAY(prog_info_const, each_case)
			PERSON_ARRAY(ref_numb_const, pers_count) = ref_numb
			PERSON_ARRAY(MFIP_memb_code_const, pers_count) = member_code

			row = row + 1
			pers_count = pers_count + 1
			EMReadScreen next_ref_numb, 2, row, 6
			If row = 18 then
				PF8
				row = 7
			End if
		Loop until next_ref_numb = "  "
	Else
		call navigate_to_MAXIS_screen("STAT", "MEMB")
		Do
			EMReadScreen ref_numb, 2, 4, 33

			ReDim preserve PERSON_ARRAY(last_const, pers_count)
			PERSON_ARRAY(case_numb_const, pers_count) = CASE_NUMBER_ARRAY(case_numb_const, each_case)
			PERSON_ARRAY(worker_numb_const, pers_count) = CASE_NUMBER_ARRAY(worker_numb_const, each_case)
			PERSON_ARRAY(prog_info_const, pers_count) = CASE_NUMBER_ARRAY(prog_info_const, each_case)
			PERSON_ARRAY(ref_numb_const, pers_count) = ref_numb
			PERSON_ARRAY(MFIP_memb_code_const, pers_count) = "PEND"

			pers_count = pers_count + 1
			transmit
			EMReadScreen last_page, 7, 24, 2
		Loop until last_page = "ENTER A"
	End If


	Call back_to_SELF
	call navigate_to_MAXIS_screen("STAT", "MEMB")
	EMWaitReady 0, 0
	For each_pers = first_pers_index to UBound(PERSON_ARRAY, 2)
		Call write_value_and_transmit(PERSON_ARRAY(ref_numb_const, each_pers), 20, 76)

		EMReadScreen age, 3, 8, 76
		EMReadScreen pmi_numb, 8, 4, 46
		PERSON_ARRAY(age_const, each_pers) = trim(age)
		If PERSON_ARRAY(age_const, each_pers) = "" Then PERSON_ARRAY(age_const, each_pers) = 0
		PERSON_ARRAY(pmi_numb_const, each_pers) = trim(pmi_numb)
		age = ""
		pmi_numb = ""
	Next

	call navigate_to_MAXIS_screen("STAT", "MEMI")
	EMWaitReady 0, 0
	For each_pers = first_pers_index to UBound(PERSON_ARRAY, 2)
		Call write_value_and_transmit(PERSON_ARRAY(ref_numb_const, each_pers), 20, 76)

		EMReadScreen mariage_code, 1, 7, 40
		PERSON_ARRAY(marital_const, each_pers) = mariage_code
	Next

	call navigate_to_MAXIS_screen("STAT", "UNEA")
	EMWaitReady 0, 0
	For each_pers = first_pers_index to UBound(PERSON_ARRAY, 2)
		Call write_value_and_transmit(PERSON_ARRAY(ref_numb_const, each_pers), 20, 76)

		Do
			EmReadScreen lost_memb, 20, 24, 15
			If lost_memb = "NOT IN THE HOUSEHOLD" Then Exit Do
			EMReadScreen income_type, 2, 5, 37

			If income_type = "03" Then
				EMReadscreen ssi_amount, 8, 18, 68
				PERSON_ARRAY(ssi_amt_const, each_pers) = trim(ssi_amount)
			ElseIf income_type = "01" or income_type = "02" Then
				EMReadscreen rsdi_amount, 8, 18, 68
				If PERSON_ARRAY(rsdi_1_amt_const, each_pers) <> "" Then PERSON_ARRAY(rsdi_2_amt_const, each_pers) = trim(rsdi_amount)
				If PERSON_ARRAY(rsdi_1_amt_const, each_pers) = "" Then PERSON_ARRAY(rsdi_1_amt_const, each_pers) = trim(rsdi_amount)
			End If

			transmit
			EMReadScreen last_page, 7, 24, 2
		Loop until last_page = "ENTER A"
	Next

	For each_pers = first_pers_index to UBound(PERSON_ARRAY, 2)
		objExcel.cells(excel_row, 1).value = PERSON_ARRAY(worker_numb_const, each_pers)
		objExcel.cells(excel_row, 2).value = PERSON_ARRAY(case_numb_const, each_pers)
		objExcel.cells(excel_row, 3).value = PERSON_ARRAY(pmi_numb_const, each_pers)
		objExcel.cells(excel_row, 4).value = PERSON_ARRAY(ref_numb_const, each_pers)
		objExcel.cells(excel_row, 5).value = PERSON_ARRAY(age_const, each_pers)
		objExcel.cells(excel_row, 6).value = PERSON_ARRAY(marital_const, each_pers)
		objExcel.cells(excel_row, 7).value = PERSON_ARRAY(MFIP_memb_code_const, each_pers)
		objExcel.cells(excel_row, 8).value = PERSON_ARRAY(rsdi_1_amt_const, each_pers)
		objExcel.cells(excel_row, 9).value = PERSON_ARRAY(rsdi_2_amt_const, each_pers)
		objExcel.cells(excel_row, 10).value = PERSON_ARRAY(ssi_amt_const, each_pers)
		objExcel.cells(excel_row, 11).value = PERSON_ARRAY(prog_info_const, each_pers)
		excel_row = excel_row + 1
	Next

Next

Call script_end_procedure("DONE")
