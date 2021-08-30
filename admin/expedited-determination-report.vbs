'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EXPEDITED DETERMINATION REPORT.vbs"
start_time = timer
STATS_counter = 0			     'sets the stats counter at one
STATS_manualtime = 	60			 'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("10/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DECLARATIONS BLOCK ========================================================================================================
const case_number_col_const 				= 1
const worker_col_const 						= 2
const xnumber_col_const 					= 3
const date_of_appl_col_const 				= 4
const appt_notc_date_col_const				= 5
const date_of_appt_col_const				= 6
const date_of_intve_col_const 				= 7
const screen_status_col_const 				= 8
const det_status_col_const 					= 9
const det_income_col_const 					= 10
const det_asset_col_const 					= 11
const det_shel_col_const 					= 12
const det_hest_col_const 					= 13
const date_of_appr_col_const 				= 14
const date_of_deny_col_const 				= 15
const deny_reason_col_const 				= 16
const id_on_file_col_const 					= 17
const outstate_action_col_const 			= 18
const outstate_state_col_const 				= 19
const outstate_end_date_rept_col_const 		= 20
const outstate_openended_col_const 			= 21
const outstate_end_date_verif_col_const 	= 22
const mn_elig_begin_col_const 				= 23
const prev_post_delay_col_const 			= 24
const prev_post_prev_date_of_appl_col_const = 25
const prev_post_list_col_const 				= 26
const prev_post_curr_verif_post_col_const 	= 27
const prev_post_reg_snap_app_col_const 		= 28
const prev_post_verifs_recvd_col_const 		= 29
const expl_appr_delay_col_const 			= 30
const post_verifs_yn_col_const 				= 31
const post_verifs_list_col_const 			= 32
const faci_delay_col_const 					= 33
const faci_deny_col_const 					= 34
const faci_name_col_const 					= 35
const faci_snap_inelig_col_const 			= 36
const faci_entry_col_const 					= 37
const faci_release_col_const 				= 38
const faci_release_in_30_col_const 			= 39
const script_run_date_col_const 			= 40
'END DECLARATIONS BLOCK ====================================================================================================

'There is no EMConnect and no MAXIS checking because this script does not use MAXIS at all
'Declaring the only dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 90, "Expedited Determination Report"
  Text 10, 10, 225, 20, "This script will read the tracking TXT files from the Assignments foler and add the information into the Expedited Determination Report. "
  Text 10, 35, 225, 10, "When the script is complete, the Excel will be saved."
  Text 10, 50, 130, 20, "At the end of the script run, would you like the Excel file to remain open:"
  DropListBox 10, 70, 180, 45, "No - Close the file"+chr(9)+"Yes - Keep the file open", leave_excel_open
  ButtonGroup ButtonPressed
    OkButton 200, 70, 40, 15
EndDialog

'showing the dialog - there is no loop because there is nothing to manage and no password handling.
dialog Dialog1
cancel_confirmation

'defining the assignment folder
exp_assignment_folder = t_drive & "\Eligibility Support\Assignments\Expedited Information"
Set objFolder = objFSO.GetFolder(exp_assignment_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

'Open an existing Excel for the year
report_out_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\2021 EXP Determination Report Out.xlsx"

Call excel_open(report_out_file, True, True, ObjExcel, objWorkbook)  			'opens the selected excel file'

For Each objWorkSheet In objWorkbook.Worksheets									'looking through each of the worksheets to find the 'ALL CASES' worksheet
    If instr(objWorkSheet.Name, "ALL CASES") <> 0 Then
		set objALLCASESWorkSheet = objWorkSheet									'setting the 'ALL CASES' to a worksheet variable because we need it a lot
        objALLCASESWorkSheet.Activate											'opening that worksheet
        Exit For
    End If
Next

'Now we need to find the last row in the 'ALL CASES' sheet so we don't overwrite anything
total_excel_row = 1																'default to the first row
Do
    total_excel_row = total_excel_row + 1
    this_case_number = trim(ObjExcel.Cells(total_excel_row, 1).Value)
Loop Until this_case_number = ""												'if the case number is blank then the row is blank

ObjExcel.Columns(deny_reason_col_const).ColumnWidth = 150						'setting some column widths
ObjExcel.Columns(deny_reason_col_const).WrapText = True
ObjExcel.Columns(expl_appr_delay_col_const).ColumnWidth = 150
ObjExcel.Columns(expl_appr_delay_col_const).WrapText = True

'Add a sheet to the Excel with the report date
sheet_friendly_date = replace(date, "/", "-")
sheet_name = sheet_friendly_date & " REPT"
ObjExcel.Worksheets.Add().Name = sheet_name

For Each objWorkSheet In objWorkbook.Worksheets									'setting the worksheet to a variable so we can use it again
    If objWorkSheet.Name = sheet_name Then
		set objTODAYWorkSheet = objWorkSheet
		objTODAYWorkSheet.Activate
        Exit For
    End If
Next

'ADD HEADERS HERE - to the new sheet
ObjExcel.Cells(1, case_number_col_const).Value  				= "CASE NUMBER"
ObjExcel.Cells(1, worker_col_const).Value  						= "WORKER NAME"
ObjExcel.Cells(1, xnumber_col_const).Value  					= "CASE X NUMBER"
ObjExcel.Cells(1, date_of_appl_col_const).Value  				= "DATE OF APPLICATION"
ObjExcel.Cells(1, appt_notc_date_col_const).Value  				= "APPT NOTC SENT DATE"
ObjExcel.Cells(1, date_of_appt_col_const).Value  				= "DATE OF APPT"
ObjExcel.Cells(1, date_of_intve_col_const).Value  				= "DATE OF INTERVIEW"
ObjExcel.Cells(1, screen_status_col_const).Value  				= "EXPEDITED SCREENING STATUS"
ObjExcel.Cells(1, det_status_col_const).Value  					= "EXPEDITED DETERMINATION STATUS"
ObjExcel.Cells(1, det_income_col_const).Value 					= "INCOME"
ObjExcel.Cells(1, det_asset_col_const).Value 					= "ASSETS"
ObjExcel.Cells(1, det_shel_col_const).Value 					= "SHELTER"
ObjExcel.Cells(1, det_hest_col_const).Value 					= "UTILITIES"
ObjExcel.Cells(1, date_of_appr_col_const).Value  				= "DATE OF APPROVAL"
ObjExcel.Cells(1, date_of_deny_col_const).Value  				= "SNAP DENIAL DATE"
ObjExcel.Cells(1, deny_reason_col_const).Value 					= "SNAP DENIAL REASON"
ObjExcel.Cells(1, id_on_file_col_const).Value 					= "ID ON FILE"
ObjExcel.Cells(1, outstate_action_col_const).Value 				= "OUT STATE ACTION"
ObjExcel.Cells(1, outstate_state_col_const).Value 				= "OUT STATE STATE"
ObjExcel.Cells(1, outstate_end_date_rept_col_const).Value 		= "OUT STATE REPORTED END"
ObjExcel.Cells(1, outstate_openended_col_const).Value 			= "OUT STATE OPEN ENDED"
ObjExcel.Cells(1, outstate_end_date_verif_col_const).Value 		= "OUT STATE VERIFIED END"
ObjExcel.Cells(1, mn_elig_begin_col_const).Value 				= "MN ELIG BEGIN"
ObjExcel.Cells(1, prev_post_delay_col_const).Value 				= "PREV POSTPND CAUSE DELAY"
ObjExcel.Cells(1, prev_post_prev_date_of_appl_col_const).Value 	= "PREV POSTPND PREV DATE OF APPL"
ObjExcel.Cells(1, prev_post_list_col_const).Value 				= "PREV POSTPND LIST"
ObjExcel.Cells(1, prev_post_curr_verif_post_col_const).Value 	= "PREV POSTPND CURR VERIF POST"
ObjExcel.Cells(1, prev_post_reg_snap_app_col_const).Value 		= "PREV POSTPND REG SNAP APPR"
ObjExcel.Cells(1, prev_post_verifs_recvd_col_const).Value 		= "PREV POSTPND VERIFS RECVD"
ObjExcel.Cells(1, expl_appr_delay_col_const).Value 				= "EXPLAIN APPROVAL DELAYS "
ObjExcel.Cells(1, post_verifs_yn_col_const).Value 				= "POSTPONED VERIFICATIONS"
ObjExcel.Cells(1, post_verifs_list_col_const).Value 			= "WHAT ARE THE POSTPONED VERIFICATIONS"
ObjExcel.Cells(1, faci_delay_col_const).Value 					= "FACI CASUE DELAY"
ObjExcel.Cells(1, faci_deny_col_const).Value 					= "FACI CAUSE DENY"
ObjExcel.Cells(1, faci_name_col_const).Value 					= "FACI NAME"
ObjExcel.Cells(1, faci_snap_inelig_col_const).Value 			= "FACI INELIG SNAP"
ObjExcel.Cells(1, faci_entry_col_const).Value 					= "FACI ENTRY"
ObjExcel.Cells(1, faci_release_col_const).Value 				= "FACI RELEASE"
ObjExcel.Cells(1, faci_release_in_30_col_const).Value 			= "FACI RELEASE IN 30"
ObjExcel.Cells(1, script_run_date_col_const).Value 				= "DATE OF SCRIPT RUN"
ObjExcel.Rows(1).Font.Bold = True

excel_row = 2		'setting the first row

For Each objFile in colFiles																'looping through each file

    this_file_path = objFile.Path
    'Setting the object to open the text file for reading the data already in the file
    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

    'Reading the entire text file into a string
    every_line_in_text_file = objTextStream.ReadAll

    exp_det_details = split(every_line_in_text_file, vbNewLine)					'creating an array of all of the information in the TXT files

	objALLCASESWorkSheet.Activate												'go to the ALL CASES sheet
    For Each text_line in exp_det_details										'read each line in the file
        If Instr(text_line, "^*^*^") <> 0 Then
            line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
            line_info(0) = trim(line_info(0))
			'here we add the information from TXT to Excel
            If line_info(0) = "CASE NUMBER"                             Then ObjExcel.Cells(total_excel_row, case_number_col_const).Value  = line_info(1)
            If line_info(0) = "WORKER NAME"                             Then ObjExcel.Cells(total_excel_row, worker_col_const).Value  = line_info(1)
            If line_info(0) = "CASE X NUMBER"                           Then ObjExcel.Cells(total_excel_row, xnumber_col_const).Value  = line_info(1)
            If line_info(0) = "DATE OF APPLICATION"                     Then ObjExcel.Cells(total_excel_row, date_of_appl_col_const).Value  = line_info(1)
			If line_info(0) = "APPT NOTC SENT DATE"                     Then ObjExcel.Cells(total_excel_row, appt_notc_date_col_const).Value  = line_info(1)
			If line_info(0) = "APPT DATE"                     			Then ObjExcel.Cells(total_excel_row, date_of_appt_col_const).Value  = line_info(1)
            If line_info(0) = "DATE OF INTERVIEW"                       Then ObjExcel.Cells(total_excel_row, date_of_intve_col_const).Value  = line_info(1)
            If line_info(0) = "EXPEDITED SCREENING STATUS"              Then ObjExcel.Cells(total_excel_row, screen_status_col_const).Value  = line_info(1)
            If line_info(0) = "EXPEDITED DETERMINATION STATUS"          Then ObjExcel.Cells(total_excel_row, det_status_col_const).Value  = line_info(1)
			If line_info(0) = "DET INCOME" 								Then ObjExcel.Cells(total_excel_row, det_income_col_const).Value  = line_info(1)
			If line_info(0) = "DET ASSETS" 								Then ObjExcel.Cells(total_excel_row, det_asset_col_const).Value  = line_info(1)
			If line_info(0) = "DET SHEL" 								Then ObjExcel.Cells(total_excel_row, det_shel_col_const).Value  = line_info(1)
			If line_info(0) = "DET HEST" 								Then ObjExcel.Cells(total_excel_row, det_hest_col_const).Value  = line_info(1)
            If line_info(0) = "DATE OF APPROVAL"                        Then ObjExcel.Cells(total_excel_row, date_of_appr_col_const).Value  = line_info(1)
            If line_info(0) = "SNAP DENIAL DATE"                        Then ObjExcel.Cells(total_excel_row, date_of_deny_col_const).Value  = line_info(1)
            If line_info(0) = "SNAP DENIAL REASON"                      Then ObjExcel.Cells(total_excel_row, deny_reason_col_const).Value = line_info(1)
            If line_info(0) = "ID ON FILE"                              Then ObjExcel.Cells(total_excel_row, id_on_file_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE ACTION" 						Then ObjExcel.Cells(total_excel_row, outstate_action_col_const).Value  = line_info(1)
			If line_info(0) = "OUTSTATE STATE" 							Then ObjExcel.Cells(total_excel_row, outstate_state_col_const).Value  = line_info(1)
            If line_info(0) = "END DATE OF SNAP IN ANOTHER STATE"       Then ObjExcel.Cells(total_excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE REPORTED END DATE"				Then ObjExcel.Cells(total_excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE OPENENDED" 						Then ObjExcel.Cells(total_excel_row, outstate_openended_col_const).Value  = line_info(1)
			If line_info(0) = "OUTSTATE VERIFIED END DATE" 				Then ObjExcel.Cells(total_excel_row, outstate_end_date_verif_col_const).Value  = line_info(1)
			If line_info(0) = "MN ELIG BEGIN DATE" 						Then ObjExcel.Cells(total_excel_row, mn_elig_begin_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST DELAY APP" 					Then ObjExcel.Cells(total_excel_row, prev_post_delay_col_const).Value = line_info(1)
			If line_info(0) = "EXPEDITED APPROVE PREVIOUSLY POSTPONED"  Then ObjExcel.Cells(total_excel_row, prev_post_delay_col_const).Value = line_info(1)
			If line_info(0) = "PREV POST PREV DATE OF APP" 				Then ObjExcel.Cells(total_excel_row, prev_post_prev_date_of_appl_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST LIST" 							Then ObjExcel.Cells(total_excel_row, prev_post_list_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST CURR VERIF POST" 				Then ObjExcel.Cells(total_excel_row, prev_post_curr_verif_post_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST ONGOING SNAP APP" 				Then ObjExcel.Cells(total_excel_row, prev_post_reg_snap_app_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST VERIFS RECVD" 					Then ObjExcel.Cells(total_excel_row, prev_post_verifs_recvd_col_const).Value  = line_info(1)
            If line_info(0) = "EXPLAIN APPROVAL DELAYS"                 Then ObjExcel.Cells(total_excel_row, expl_appr_delay_col_const).Value = line_info(1)
            If line_info(0) = "POSTPONED VERIFICATIONS"                 Then ObjExcel.Cells(total_excel_row, post_verifs_yn_col_const).Value = line_info(1)
            If line_info(0) = "WHAT ARE THE POSTPONED VERIFICATIONS"    Then ObjExcel.Cells(total_excel_row, post_verifs_list_col_const).Value = line_info(1)
			If line_info(0) = "FACI DELAY ACTION" 						Then ObjExcel.Cells(total_excel_row, faci_delay_col_const).Value  = line_info(1)
			If line_info(0) = "FACI DENY" 								Then ObjExcel.Cells(total_excel_row, faci_deny_col_const).Value  = line_info(1)
			If line_info(0) = "FACI NAME" 								Then ObjExcel.Cells(total_excel_row, faci_name_col_const).Value  = line_info(1)
			If line_info(0) = "FACI INELIG SNAP" 						Then ObjExcel.Cells(total_excel_row, faci_snap_inelig_col_const).Value  = line_info(1)
			If line_info(0) = "FACI ENTRY DATE" 						Then ObjExcel.Cells(total_excel_row, faci_entry_col_const).Value  = line_info(1)
			If line_info(0) = "FACI RELEASE DATE" 						Then ObjExcel.Cells(total_excel_row, faci_release_col_const).Value  = line_info(1)
			If line_info(0) = "FACI RELEASE IN 30 DAYS" 				Then ObjExcel.Cells(total_excel_row, faci_release_in_30_col_const).Value  = line_info(1)
            If line_info(0) = "DATE OF SCRIPT RUN"                      Then ObjExcel.Cells(total_excel_row, script_run_date_col_const).Value = line_info(1)
        End If
    Next
    total_excel_row = total_excel_row + 1										'advance to the next row

	objTODAYWorkSheet.Activate													'opening the daily sheet
	For Each text_line in exp_det_details										'read each line in the file
		If Instr(text_line, "^*^*^") <> 0 Then
			line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
			line_info(0) = trim(line_info(0))
			'here we add the information from TXT to Excel
			If line_info(0) = "CASE NUMBER"                             Then ObjExcel.Cells(excel_row, case_number_col_const).Value  = line_info(1)
			If line_info(0) = "WORKER NAME"                             Then ObjExcel.Cells(excel_row, worker_col_const).Value  = line_info(1)
			If line_info(0) = "CASE X NUMBER"                           Then ObjExcel.Cells(excel_row, xnumber_col_const).Value  = line_info(1)
			If line_info(0) = "DATE OF APPLICATION"                     Then ObjExcel.Cells(excel_row, date_of_appl_col_const).Value  = line_info(1)
			If line_info(0) = "APPT NOTC SENT DATE"                     Then ObjExcel.Cells(excel_row, appt_notc_date_col_const).Value  = line_info(1)
			If line_info(0) = "APPT DATE"                     			Then ObjExcel.Cells(excel_row, date_of_appt_col_const).Value  = line_info(1)
			If line_info(0) = "DATE OF INTERVIEW"                       Then ObjExcel.Cells(excel_row, date_of_intve_col_const).Value  = line_info(1)
			If line_info(0) = "EXPEDITED SCREENING STATUS"              Then ObjExcel.Cells(excel_row, screen_status_col_const).Value  = line_info(1)
			If line_info(0) = "EXPEDITED DETERMINATION STATUS"          Then ObjExcel.Cells(excel_row, det_status_col_const).Value  = line_info(1)
			If line_info(0) = "DET INCOME" 								Then ObjExcel.Cells(excel_row, det_income_col_const).Value  = line_info(1)
			If line_info(0) = "DET ASSETS" 								Then ObjExcel.Cells(excel_row, det_asset_col_const).Value  = line_info(1)
			If line_info(0) = "DET SHEL" 								Then ObjExcel.Cells(excel_row, det_shel_col_const).Value  = line_info(1)
			If line_info(0) = "DET HEST" 								Then ObjExcel.Cells(excel_row, det_hest_col_const).Value  = line_info(1)
			If line_info(0) = "DATE OF APPROVAL"                        Then ObjExcel.Cells(excel_row, date_of_appr_col_const).Value  = line_info(1)
			If line_info(0) = "SNAP DENIAL DATE"                        Then ObjExcel.Cells(excel_row, date_of_deny_col_const).Value  = line_info(1)
			If line_info(0) = "SNAP DENIAL REASON"                      Then ObjExcel.Cells(excel_row, deny_reason_col_const).Value = line_info(1)
			If line_info(0) = "ID ON FILE"                              Then ObjExcel.Cells(excel_row, id_on_file_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE ACTION" 						Then ObjExcel.Cells(excel_row, outstate_action_col_const).Value  = line_info(1)
			If line_info(0) = "OUTSTATE STATE" 							Then ObjExcel.Cells(excel_row, outstate_state_col_const).Value  = line_info(1)
			If line_info(0) = "END DATE OF SNAP IN ANOTHER STATE"       Then ObjExcel.Cells(excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE REPORTED END DATE"				Then ObjExcel.Cells(excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
			If line_info(0) = "OUTSTATE OPENENDED" 						Then ObjExcel.Cells(excel_row, outstate_openended_col_const).Value  = line_info(1)
			If line_info(0) = "OUTSTATE VERIFIED END DATE" 				Then ObjExcel.Cells(excel_row, outstate_end_date_verif_col_const).Value  = line_info(1)
			If line_info(0) = "MN ELIG BEGIN DATE" 						Then ObjExcel.Cells(excel_row, mn_elig_begin_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST DELAY APP" 					Then ObjExcel.Cells(excel_row, prev_post_delay_col_const).Value = line_info(1)
			If line_info(0) = "EXPEDITED APPROVE PREVIOUSLY POSTPONED"  Then ObjExcel.Cells(excel_row, prev_post_delay_col_const).Value = line_info(1)
			If line_info(0) = "PREV POST PREV DATE OF APP" 				Then ObjExcel.Cells(excel_row, prev_post_prev_date_of_appl_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST LIST" 							Then ObjExcel.Cells(excel_row, prev_post_list_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST CURR VERIF POST" 				Then ObjExcel.Cells(excel_row, prev_post_curr_verif_post_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST ONGOING SNAP APP" 				Then ObjExcel.Cells(excel_row, prev_post_reg_snap_app_col_const).Value  = line_info(1)
			If line_info(0) = "PREV POST VERIFS RECVD" 					Then ObjExcel.Cells(excel_row, prev_post_verifs_recvd_col_const).Value  = line_info(1)
			If line_info(0) = "EXPLAIN APPROVAL DELAYS"                 Then ObjExcel.Cells(excel_row, expl_appr_delay_col_const).Value = line_info(1)
			If line_info(0) = "POSTPONED VERIFICATIONS"                 Then ObjExcel.Cells(excel_row, post_verifs_yn_col_const).Value = line_info(1)
			If line_info(0) = "WHAT ARE THE POSTPONED VERIFICATIONS"    Then ObjExcel.Cells(excel_row, post_verifs_list_col_const).Value = line_info(1)
			If line_info(0) = "FACI DELAY ACTION" 						Then ObjExcel.Cells(excel_row, faci_delay_col_const).Value  = line_info(1)
			If line_info(0) = "FACI DENY" 								Then ObjExcel.Cells(excel_row, faci_deny_col_const).Value  = line_info(1)
			If line_info(0) = "FACI NAME" 								Then ObjExcel.Cells(excel_row, faci_name_col_const).Value  = line_info(1)
			If line_info(0) = "FACI INELIG SNAP" 						Then ObjExcel.Cells(excel_row, faci_snap_inelig_col_const).Value  = line_info(1)
			If line_info(0) = "FACI ENTRY DATE" 						Then ObjExcel.Cells(excel_row, faci_entry_col_const).Value  = line_info(1)
			If line_info(0) = "FACI RELEASE DATE" 						Then ObjExcel.Cells(excel_row, faci_release_col_const).Value  = line_info(1)
			If line_info(0) = "FACI RELEASE IN 30 DAYS" 				Then ObjExcel.Cells(excel_row, faci_release_in_30_col_const).Value  = line_info(1)
			If line_info(0) = "DATE OF SCRIPT RUN"                      Then ObjExcel.Cells(excel_row, script_run_date_col_const).Value = line_info(1)
		End If
	Next
	excel_row = excel_row + 1													'advance to the next row

	STATS_counter = STATS_counter + 1
	objTextStream.Close						'we are done with this file, so we must close the access
    objFSO.DeleteFile(this_file_path)		'deleting the TXT file because hgave the information
Next

'formatting the worksheet made for the day
objTODAYWorkSheet.Activate
Const xlSrcRange = 1
Const xlYes = 1
xlVAlignTop = -4160
xlHAlignLeft = -4131
For col = 1 to 17
    ObjExcel.columns(col).AutoFit()
    ObjExcel.columns(col).VerticalAlignment = xlVAlignTop
    ObjExcel.columns(col).HorizontalAlignment = xlHAlignLeft
Next

'setting some column widths on the day sheet
ObjExcel.Columns(deny_reason_col_const).ColumnWidth = 150
ObjExcel.Columns(deny_reason_col_const).WrapText = True
ObjExcel.Columns(expl_appr_delay_col_const).ColumnWidth = 150
ObjExcel.Columns(expl_appr_delay_col_const).WrapText = True

'here we add the table format to the sheet from today
tableRange = "A1:AN" & excel_row-1
table_friendly_date = replace(date, "/", "")
table_friendly_date = trim(table_friendly_date)
table_name = table_friendly_date & "TABLE"
ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, tableRange, xlYes).Name = table_name
' ObjExcel.ActiveSheet.ListObjects(table_name).TableStyle = "TableStyleDark2"

'ending on the statistics sheet
For Each objWorkSheet In objWorkbook.Worksheets
    If objWorkSheet.Name = "Statistics" Then
		objWorkSheet.Activate
        Exit For
    End If
Next

objWorkbook.Save()		'saving the excel

If leave_excel_open = "No - Close the file" Then		'if the file should be closed - it does it here.
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit
End If

'SAVE EXCEL'
Call script_end_procedure("Expedited Determination report is updated and the tracking files removed.")
