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

ex_parte_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte"
ep_revw_mo = ""

'define excel columns
case_number_column = 1
smi_column = 5
UserName = "daco003"
year_month = "202406"
date_time = now
ep_revw_mo = "09"
ep_revw_yr = "24"

Const adOpenStatic = 3
Const adLockOptimistic = 3

'Open up the excel file
Call excel_open(ex_parte_folder & "\AVS Lists\AVS List - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx", True, False, objAVSExcel, objAVSworkbook)
AVS_excel_row = 4
'Find first row with case number

'open table connection
    Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
Do
    memb_smi = ObjAVSExcel.Cells(avs_excel_row, smi_column)
    MAXIS_case_number = ObjAVSExcel.Cells(avs_excel_row, case_number_column)
    asset_test = 1
    
    'insert into table
    objRecordSet.Open "INSERT INTO ES.ES_AVSList (YearMonth, SMI, CaseNumber, AssetTest)" &  _
			"VALUES ('" & year_month & "', '" & memb_smi & "', '" & MAXIS_case_number & "', '" & asset_test & "')", objConnection, adOpenStatic, adLockOptimistic
    avs_excel_row = AVS_excel_row + 1


Loop until ObjAvsExcel.Cells(avs_excel_row, case_number_column) = ""

'close connection

objConnection.close

msgbox "Complete. "  & avs_excel_row & " records entered to DB"
stopscript

