'Required for statistical purposes==========================================================================================
name_of_script = "GRH Budg Type.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================
run_locally = TRUE
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

' file_url = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\QI On Demand Daily Assignment\QI On Demand Work Log.xlsx"
' "https://hennepin-my.sharepoint.com/:x:/g/personal/casey_love_hennepin_us/EcnENuYdvEZKi9ki9FI0s3UB140-usJKZBY-Rcg23AMk0Q?email=Casey.Love%40hennepin.us&e=bN0HwR"
' "https://hennepin-my.sharepoint.com/:x:/g/personal/casey_love_hennepin_us/EcnENuYdvEZKi9ki9FI0s3UB140-usJKZBY-Rcg23AMk0Q?email=Casey.Love%40hennepin.us&e=Rdc4qP"
file_url = "C:\Users\" & user_ID_for_validation & "OneDrive - Hennepin County\TEMP Files\ACTV 8-13-24.xlsx"
file_url = user_myDocs_folder & "ACTV 8-13-24.xlsx"
Call excel_open(file_url, True, False, ObjExcel, objWorkbook)
MAXIS_footer_month = "10"
MAXIS_footer_year = "24"

Worker_col = 1
case_numb_col = 2
grh_status_col = 9
budget_type_col = 12

excel_row = 2
Do While trim(ObjExcel.Cells(excel_row, Worker_col).value) <> ""
	If trim(ObjExcel.Cells(excel_row, grh_status_col).value) = "A" Then
		MAXIS_case_number = trim(ObjExcel.Cells(excel_row, case_numb_col).value)
		Call back_to_SELF
		Call navigate_to_MAXIS_screen("ELIG", "GRH ")
		Call write_value_and_transmit("GRPB", 20, 71)

		EMReadScreen Budg_info, 33, 3, 23
		Budg_info = trim(Budg_info)
		ObjExcel.Cells(excel_row, budget_type_col).value = Budg_info
	End If
	excel_row = excel_row + 1
Loop

Call script_end_procedure ("Done?")