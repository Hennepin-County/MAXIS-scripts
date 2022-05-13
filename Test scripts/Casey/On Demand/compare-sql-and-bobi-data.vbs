'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
' call run_from_GitHub(script_repository & "application-received.vbs")

'END FUNCTIONS LIBRARY BLOCK================================================================================================
EMConnect""

const case_numb = 0
const app_date  = 1
const found     = 2
const last_const = 3

Dim BOBI_LIST_ARRAY()
ReDim BOBI_LIST_ARRAY(last_const, 0)

Dim SQL_LIST_ARRAY()
ReDim SQL_LIST_ARRAY(last_const, 0)

bobi_cases_string = "~"
sql_cases_string = "~"


'Open the days excel
excel_name = "5-13-2022 Data Review.xlsx"
excel_path = "C:\Users\calo001\OneDrive - Hennepin County\Projects\On Demand\BOBI vs SQL Data Review\" & excel_name
Call excel_open(excel_path, True, False, ObjExcel, objWorkbook)

'read the BOBI List
ObjExcel.worksheets("BOBI List").Activate

excel_row = 2
array_item = 0
Do
    bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
    If InStr(bobi_cases_string, "~" & bobi_case_numb & "~") = 0 Then
        ReDim Preserve BOBI_LIST_ARRAY(last_const, array_item)
        BOBI_LIST_ARRAY(case_numb, array_item) = bobi_case_numb
        BOBI_LIST_ARRAY(app_date, array_item) = ObjExcel.Cells(excel_row, 6).Value

        array_item = array_item + 1
    End If
    excel_row = excel_row + 1
    next_bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
Loop until next_bobi_case_numb = ""

'read the SQL list
ObjExcel.worksheets("SQL List").Activate

excel_row = 2
array_item = 0
Do
    bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
    If InStr(bobi_cases_string, "~" & bobi_case_numb & "~") = 0 Then
        ReDim Preserve SQL_LIST_ARRAY(last_const, array_item)
        SQL_LIST_ARRAY(case_numb, array_item) = bobi_case_numb
        SQL_LIST_ARRAY(app_date, array_item) = ObjExcel.Cells(excel_row, 4).Value
        SQL_LIST_ARRAY(found, array_item) = False

        array_item = array_item + 1
    End If
    excel_row = excel_row + 1
    next_bobi_case_numb = trim(ObjExcel.Cells(excel_row, 2).Value)
Loop until next_bobi_case_numb = ""

'loop through the BOBI list add the SQL
ObjExcel.worksheets("Data Compare").Activate

excel_row = 2
For bobi_item = 0 to UBound(BOBI_LIST_ARRAY, 2)
    MAXIS_case_number = BOBI_LIST_ARRAY(case_numb, bobi_item)
    ObjExcel.Cells(excel_row, 1).Value = BOBI_LIST_ARRAY(case_numb, bobi_item)
    ObjExcel.Cells(excel_row, 2).Value = True
    ObjExcel.Cells(excel_row, 3).Value = BOBI_LIST_ARRAY(app_date, bobi_item)
    ObjExcel.Cells(excel_row, 4).Value = False
    For sql_item = 0 to UBound(SQL_LIST_ARRAY, 2)
        If SQL_LIST_ARRAY(case_numb, sql_item) = BOBI_LIST_ARRAY(case_numb, bobi_item) Then
            SQL_LIST_ARRAY(found, sql_item) = True
            ObjExcel.Cells(excel_row, 4).Value = True
            ObjExcel.Cells(excel_row, 5).Value = SQL_LIST_ARRAY(app_date, sql_item)
            Exit For
        End If
    Next
    Call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen cash_1_status, 4, 6, 74
    EMReadScreen cash_1_intvw, 8, 6, 55
    EMReadScreen cash_2_status, 4, 7, 74
    EMReadScreen cash_2_intvw, 8, 7, 55
    EMReadScreen snap_status, 4, 10, 74
    EMReadScreen snap_intvw, 8, 10, 55

    cash_1_intvw = replace(cash_1_intvw, " ", "/")
    If cash_1_intvw = "__/__/__" Then cash_1_intvw = ""
    cash_2_intvw = replace(cash_2_intvw, " ", "/")
    If cash_2_intvw = "__/__/__" Then cash_2_intvw = ""
    snap_intvw = replace(snap_intvw, " ", "/")
    If snap_intvw = "__/__/__" Then snap_intvw = ""

    ObjExcel.Cells(excel_row, 6).Value = cash_1_status
    ObjExcel.Cells(excel_row, 7).Value = cash_1_intvw
    ObjExcel.Cells(excel_row, 8).Value = cash_2_status
    ObjExcel.Cells(excel_row, 9).Value = cash_2_intvw
    ObjExcel.Cells(excel_row, 10).Value = snap_status
    ObjExcel.Cells(excel_row, 11).Value = snap_intvw

    Call back_to_SELF

    excel_row = excel_row + 1
Next

For sql_item = 0 to UBound(SQL_LIST_ARRAY, 2)
    If SQL_LIST_ARRAY(found, sql_item) = False Then
        MAXIS_case_number = SQL_LIST_ARRAY(case_numb, sql_item)
        ObjExcel.Cells(excel_row, 1).Value = SQL_LIST_ARRAY(case_numb, sql_item)
        ObjExcel.Cells(excel_row, 2).Value = False
        ObjExcel.Cells(excel_row, 3).Value = ""
        ObjExcel.Cells(excel_row, 4).Value = True
        ObjExcel.Cells(excel_row, 5).Value = SQL_LIST_ARRAY(app_date, sql_item)

        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen cash_1_status, 4, 6, 74
        EMReadScreen cash_1_intvw, 8, 6, 55
        EMReadScreen cash_2_status, 4, 7, 74
        EMReadScreen cash_2_intvw, 8, 7, 55
        EMReadScreen snap_status, 4, 10, 74
        EMReadScreen snap_intvw, 8, 10, 55

        cash_1_intvw = replace(cash_1_intvw, " ", "/")
        If cash_1_intvw = "__/__/__" Then cash_1_intvw = ""
        cash_2_intvw = replace(cash_2_intvw, " ", "/")
        If cash_2_intvw = "__/__/__" Then cash_2_intvw = ""
        snap_intvw = replace(snap_intvw, " ", "/")
        If snap_intvw = "__/__/__" Then snap_intvw = ""

        ObjExcel.Cells(excel_row, 6).Value = cash_1_status
        ObjExcel.Cells(excel_row, 7).Value = cash_1_intvw
        ObjExcel.Cells(excel_row, 8).Value = cash_2_status
        ObjExcel.Cells(excel_row, 9).Value = cash_2_intvw
        ObjExcel.Cells(excel_row, 10).Value = snap_status
        ObjExcel.Cells(excel_row, 11).Value = snap_intvw

        Call back_to_SELF

        excel_row = excel_row + 1
    End If
Next

call script_end_procedure("Done")
