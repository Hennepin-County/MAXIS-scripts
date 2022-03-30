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
const script_run_col_const					= 41

const disc_case_nbr_col_const 			= 1
const disc_exp_det_col_const 			= 2
const disc_app_in_5_col_const 			= 3
const disc_app_in_7_col_const 			= 4
const disc_appl_date_col_const 			= 5
const disc_bdays_appl_notc_col_const 	= 6
const disc_cdays_appl_notc_col_const 	= 7
const disc_notc_date_col_const 			= 8
const disc_bdays_notc_intv_col_const 	= 9
const disc_cdays_notc_intv_col_const 	= 10
const disc_intv_date_col_const 			= 11
const disc_app_date_col_const 			= 12
const disc_id_col_const 				= 13
const disc_app_delays_col_const 		= 14
const disc_income_col_const 			= 15
const disc_asset_col_const 				= 16
const disc_shelter_col_const 			= 17
const disc_utilities_col_const 			= 18
const disc_screening_col_const 			= 19
const disc_worker_col_const 			= 20
const disc_script_run_col_const			= 21
const disc_bdays_appl_app_col_const 	= 22
const disc_cdays_appl_app_col_const 	= 23


const work_case_nbr_col_const 			= 1
const work_worker_col_const 			= 2
const work_appl_date_col_const 			= 3
const work_notc_date_col_const 			= 4
const work_intv_date_col_const 			= 5
const work_app_date_col_const 			= 6
const work_id_col_const 				= 7
const work_app_delays_col_const 		= 8
const work_exp_det_col_const 			= 9
const work_income_col_const 			= 10
const work_asset_col_const 				= 11
const work_shelter_col_const 			= 12
const work_utilities_col_const 			= 13
const work_script_run_date_col_const	= 14

'END DECLARATIONS BLOCK ====================================================================================================
'Manually set if you want to run the testing code for creating a worklist.
'This option is only available to scriptwriters.
create_a_test_worklist = True
'If TRUE - the script will NOT delete the files that create the reports so that our data review is not changed. It also does not update any of the other reports/files
'          You should also have MAXIS ready - there may not be good password handling. Inquiry or Production is fine.
'If FALSE - the script WILL delete the files that fill in information
'           The False option is run by Laurie weekly.
If create_a_test_worklist = True Then EMConnect ""

'There is no EMConnect and no MAXIS checking because this script does not use MAXIS at all
'Declaring the only dialog
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 130, "Expedited Determination Report"
  DropListBox 10, 55, 225, 45, "Pull Data and Create Worklist"+chr(9)+"Combine Worklists", report_selection
  DropListBox 10, 110, 180, 45, "Yes - Keep the file open"+chr(9)+"No - Close the file", leave_excel_open
  ButtonGroup ButtonPressed
    OkButton 200, 110, 40, 15
  Text 10, 10, 225, 30, "This script is used to pull reports around information gathered during the Expedited Determination script runs to provide insight in how we are handling Expedited SNAP in Hennepin County"
  Text 10, 45, 155, 10, "Select which reporting option you need to run:"
  Text 10, 75, 225, 10, "When the script is complete, the Excel will be saved."
  Text 10, 90, 130, 20, "At the end of the script run, would you like the Excel file to remain open:"
EndDialog


'showing the dialog - there is no loop because there is nothing to manage and no password handling.
dialog Dialog1
cancel_confirmation

'defining the assignment folder
exp_assignment_folder = t_drive & "\Eligibility Support\Assignments\Expedited Information"
Set objFolder = objFSO.GetFolder(exp_assignment_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

'Open an existing Excel for the year
report_out_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP Determination Report Out.xlsx"
discovery_template_worklist_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Jake's Discovery\"
worklist_template_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists\"
worklist_archive_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists\Archive\"
worklist_template_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists"

discovery_template_file = discovery_template_worklist_path & "Discovery Template.xlsx"
worklist_template_file = worklist_template_path & "Worklist Template.xlsx"
worklist_review_file = worklist_template_path & "Worklist Review Report.xlsx"

date_month = DatePart("m", date)
date_day = DatePart("d", date)
date_year = DatePart("yyyy", date)
date_header = date_month & "-" & date_day & "-" & date_year
time_header = replace(time, ":", "_")
disc_file_name = "Case List from " & date_header & ".xlsx"
work_file_name = date_header & " " & time_header & " Worklist.xlsx"

daily_discovery_path = discovery_template_worklist_path & disc_file_name
current_worklist_path = worklist_template_path & work_file_name

If report_selection = "Pull Data and Create Worklist" Then

	' daily_worklist_template_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/Archive/Worklist Template.xlsx"
	' daily_discovery_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & date_header & " Worklist.xlsx"

	If create_a_test_worklist = True Then
		Call excel_open(worklist_template_file, False, False, ObjWORKExcel, objWORKWorkbook)  			'opens the selected excel file'
		ObjWORKExcel.ActiveWorkbook.SaveAs current_worklist_path
		ObjWORKExcel.worksheets("CASE LIST").Activate
	End If

	Call excel_open(report_out_file, False, False, ObjExcel, objWorkbook)  			'opens the selected excel file'
	'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW
	' Call excel_open(discovery_template_file, True, True, ObjDISCExcel, objDISCWorkbook)  			'opens the selected excel file'

	' ObjDISCExcel.ActiveWorkbook.SaveAs daily_discovery_path
	' ObjDISCExcel.worksheets("CASE LIST").Activate

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

	'Add a sheet to the Excel with the report date - REMOVING FOR NOW
	' sheet_friendly_date = replace(date, "/", "-")
	' sheet_name = sheet_friendly_date & " REPT"
	' ObjExcel.Worksheets.Add().Name = sheet_name
	'
	' For Each objWorkSheet In objWorkbook.Worksheets									'setting the worksheet to a variable so we can use it again
	'     If objWorkSheet.Name = sheet_name Then
	' 		set objTODAYWorkSheet = objWorkSheet
	' 		objTODAYWorkSheet.Activate
	'         Exit For
	'     End If
	' Next

	' 'ADD HEADERS HERE - to the new sheet
	' ObjExcel.Cells(1, case_number_col_const).Value  				= "CASE NUMBER"
	' ObjExcel.Cells(1, worker_col_const).Value  						= "WORKER NAME"
	' ObjExcel.Cells(1, xnumber_col_const).Value  					= "CASE X NUMBER"
	' ObjExcel.Cells(1, date_of_appl_col_const).Value  				= "DATE OF APPLICATION"
	' ObjExcel.Cells(1, appt_notc_date_col_const).Value  				= "APPT NOTC SENT DATE"
	' ObjExcel.Cells(1, date_of_appt_col_const).Value  				= "DATE OF APPT"
	' ObjExcel.Cells(1, date_of_intve_col_const).Value  				= "DATE OF INTERVIEW"
	' ObjExcel.Cells(1, screen_status_col_const).Value  				= "EXPEDITED SCREENING STATUS"
	' ObjExcel.Cells(1, det_status_col_const).Value  					= "EXPEDITED DETERMINATION STATUS"
	' ObjExcel.Cells(1, det_income_col_const).Value 					= "INCOME"
	' ObjExcel.Cells(1, det_asset_col_const).Value 					= "ASSETS"
	' ObjExcel.Cells(1, det_shel_col_const).Value 					= "SHELTER"
	' ObjExcel.Cells(1, det_hest_col_const).Value 					= "UTILITIES"
	' ObjExcel.Cells(1, date_of_appr_col_const).Value  				= "DATE OF APPROVAL"
	' ObjExcel.Cells(1, date_of_deny_col_const).Value  				= "SNAP DENIAL DATE"
	' ObjExcel.Cells(1, deny_reason_col_const).Value 					= "SNAP DENIAL REASON"
	' ObjExcel.Cells(1, id_on_file_col_const).Value 					= "ID ON FILE"
	' ObjExcel.Cells(1, outstate_action_col_const).Value 				= "OUT STATE ACTION"
	' ObjExcel.Cells(1, outstate_state_col_const).Value 				= "OUT STATE STATE"
	' ObjExcel.Cells(1, outstate_end_date_rept_col_const).Value 		= "OUT STATE REPORTED END"
	' ObjExcel.Cells(1, outstate_openended_col_const).Value 			= "OUT STATE OPEN ENDED"
	' ObjExcel.Cells(1, outstate_end_date_verif_col_const).Value 		= "OUT STATE VERIFIED END"
	' ObjExcel.Cells(1, mn_elig_begin_col_const).Value 				= "MN ELIG BEGIN"
	' ObjExcel.Cells(1, prev_post_delay_col_const).Value 				= "PREV POSTPND CAUSE DELAY"
	' ObjExcel.Cells(1, prev_post_prev_date_of_appl_col_const).Value 	= "PREV POSTPND PREV DATE OF APPL"
	' ObjExcel.Cells(1, prev_post_list_col_const).Value 				= "PREV POSTPND LIST"
	' ObjExcel.Cells(1, prev_post_curr_verif_post_col_const).Value 	= "PREV POSTPND CURR VERIF POST"
	' ObjExcel.Cells(1, prev_post_reg_snap_app_col_const).Value 		= "PREV POSTPND REG SNAP APPR"
	' ObjExcel.Cells(1, prev_post_verifs_recvd_col_const).Value 		= "PREV POSTPND VERIFS RECVD"
	' ObjExcel.Cells(1, expl_appr_delay_col_const).Value 				= "EXPLAIN APPROVAL DELAYS "
	' ObjExcel.Cells(1, post_verifs_yn_col_const).Value 				= "POSTPONED VERIFICATIONS"
	' ObjExcel.Cells(1, post_verifs_list_col_const).Value 			= "WHAT ARE THE POSTPONED VERIFICATIONS"
	' ObjExcel.Cells(1, faci_delay_col_const).Value 					= "FACI CASUE DELAY"
	' ObjExcel.Cells(1, faci_deny_col_const).Value 					= "FACI CAUSE DENY"
	' ObjExcel.Cells(1, faci_name_col_const).Value 					= "FACI NAME"
	' ObjExcel.Cells(1, faci_snap_inelig_col_const).Value 			= "FACI INELIG SNAP"
	' ObjExcel.Cells(1, faci_entry_col_const).Value 					= "FACI ENTRY"
	' ObjExcel.Cells(1, faci_release_col_const).Value 				= "FACI RELEASE"
	' ObjExcel.Cells(1, faci_release_in_30_col_const).Value 			= "FACI RELEASE IN 30"
	' ObjExcel.Cells(1, script_run_date_col_const).Value 				= "DATE OF SCRIPT RUN"
	' ObjExcel.Cells(1, script_run_col_const).Value 					= "SCRIPT RUN"
	'
	' ObjExcel.Rows(1).Font.Bold = True


	excel_row = 2		'setting the first row
	work_excel_row = 2

	For Each objFile in colFiles																'looping through each file
		save_to_worklist = False
		exp_det = False
		approval_date_is_date = False
		case_nbr_hold = ""
	    this_file_path = objFile.Path
	    'Setting the object to open the text file for reading the data already in the file
	    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

	    'Reading the entire text file into a string
	    every_line_in_text_file = objTextStream.ReadAll

	    exp_det_details = split(every_line_in_text_file, vbNewLine)					'creating an array of all of the information in the TXT files

		' objALLCASESWorkSheet.Activate												'go to the ALL CASES sheet - commented out because we aren't switching right now
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
	            If line_info(0) = "SCRIPT RUN"                      		Then ObjExcel.Cells(total_excel_row, script_run_col_const).Value = line_info(1)

				If line_info(0) = "EXPEDITED DETERMINATION STATUS" Then
					If UCASE(line_info(1))&"" = "TRUE" Then exp_det = True
				End If
				If line_info(0) = "DATE OF APPROVAL" Then
					If IsDate(line_info(1)) = True Then approval_date_is_date = True
				End If
				If line_info(0) = "CASE NUMBER" Then case_nbr_hold = line_info(1)
	        End If
	    Next
	    total_excel_row = total_excel_row + 1										'advance to the next row
		If create_a_test_worklist = True Then
			If exp_det = True AND approval_date_is_date = False Then save_to_worklist = True
			If exp_det = True AND approval_date_is_date = True Then
				MAXIS_case_number = case_nbr_hold
				Call back_to_SELF
				Call navigate_to_MAXIS_screen("CASE", "CURR")
				row = 1                                                 'First we will look for SNAP
			    col = 1
			    EMSearch "FS:", row, col
			    If row <> 0 Then
			        EMReadScreen fs_status, 9, row, col + 4
			        fs_status = trim(fs_status)
			        If fs_status = "PENDING" Then save_to_worklist = True
				End If
				Call back_to_SELF
				MAXIS_case_number = ""
			End If
		End If

		' REmoving the DAY specific sheet in the Master List'
		' objTODAYWorkSheet.Activate													'opening the daily sheet
		' appl_date = ""
		' notc_date = ""
		' intvw_date = ""
		' app_date = ""
		For Each text_line in exp_det_details										'read each line in the file
			If Instr(text_line, "^*^*^") <> 0 Then
				line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
				line_info(0) = trim(line_info(0))

				' REmoving the DAY specific sheet in the Master List'
				' 'here we add the information from TXT to Excel
				' If line_info(0) = "CASE NUMBER"                             Then ObjExcel.Cells(excel_row, case_number_col_const).Value  = line_info(1)
				' If line_info(0) = "WORKER NAME"                             Then ObjExcel.Cells(excel_row, worker_col_const).Value  = line_info(1)
				' If line_info(0) = "CASE X NUMBER"                           Then ObjExcel.Cells(excel_row, xnumber_col_const).Value  = line_info(1)
				' If line_info(0) = "DATE OF APPLICATION"                     Then ObjExcel.Cells(excel_row, date_of_appl_col_const).Value  = line_info(1)
				' If line_info(0) = "APPT NOTC SENT DATE"                     Then ObjExcel.Cells(excel_row, appt_notc_date_col_const).Value  = line_info(1)
				' If line_info(0) = "APPT DATE"                     			Then ObjExcel.Cells(excel_row, date_of_appt_col_const).Value  = line_info(1)
				' If line_info(0) = "DATE OF INTERVIEW"                       Then ObjExcel.Cells(excel_row, date_of_intve_col_const).Value  = line_info(1)
				' If line_info(0) = "EXPEDITED SCREENING STATUS"              Then ObjExcel.Cells(excel_row, screen_status_col_const).Value  = line_info(1)
				' If line_info(0) = "EXPEDITED DETERMINATION STATUS"          Then ObjExcel.Cells(excel_row, det_status_col_const).Value  = line_info(1)
				' If line_info(0) = "DET INCOME" 								Then ObjExcel.Cells(excel_row, det_income_col_const).Value  = line_info(1)
				' If line_info(0) = "DET ASSETS" 								Then ObjExcel.Cells(excel_row, det_asset_col_const).Value  = line_info(1)
				' If line_info(0) = "DET SHEL" 								Then ObjExcel.Cells(excel_row, det_shel_col_const).Value  = line_info(1)
				' If line_info(0) = "DET HEST" 								Then ObjExcel.Cells(excel_row, det_hest_col_const).Value  = line_info(1)
				' If line_info(0) = "DATE OF APPROVAL"                        Then ObjExcel.Cells(excel_row, date_of_appr_col_const).Value  = line_info(1)
				' If line_info(0) = "SNAP DENIAL DATE"                        Then ObjExcel.Cells(excel_row, date_of_deny_col_const).Value  = line_info(1)
				' If line_info(0) = "SNAP DENIAL REASON"                      Then ObjExcel.Cells(excel_row, deny_reason_col_const).Value = line_info(1)
				' If line_info(0) = "ID ON FILE"                              Then ObjExcel.Cells(excel_row, id_on_file_col_const).Value = line_info(1)
				' If line_info(0) = "OUTSTATE ACTION" 						Then ObjExcel.Cells(excel_row, outstate_action_col_const).Value  = line_info(1)
				' If line_info(0) = "OUTSTATE STATE" 							Then ObjExcel.Cells(excel_row, outstate_state_col_const).Value  = line_info(1)
				' If line_info(0) = "END DATE OF SNAP IN ANOTHER STATE"       Then ObjExcel.Cells(excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
				' If line_info(0) = "OUTSTATE REPORTED END DATE"				Then ObjExcel.Cells(excel_row, outstate_end_date_rept_col_const).Value = line_info(1)
				' If line_info(0) = "OUTSTATE OPENENDED" 						Then ObjExcel.Cells(excel_row, outstate_openended_col_const).Value  = line_info(1)
				' If line_info(0) = "OUTSTATE VERIFIED END DATE" 				Then ObjExcel.Cells(excel_row, outstate_end_date_verif_col_const).Value  = line_info(1)
				' If line_info(0) = "MN ELIG BEGIN DATE" 						Then ObjExcel.Cells(excel_row, mn_elig_begin_col_const).Value  = line_info(1)
				' If line_info(0) = "PREV POST DELAY APP" 					Then ObjExcel.Cells(excel_row, prev_post_delay_col_const).Value = line_info(1)
				' If line_info(0) = "EXPEDITED APPROVE PREVIOUSLY POSTPONED"  Then ObjExcel.Cells(excel_row, prev_post_delay_col_const).Value = line_info(1)
				' If line_info(0) = "PREV POST PREV DATE OF APP" 				Then ObjExcel.Cells(excel_row, prev_post_prev_date_of_appl_col_const).Value  = line_info(1)
				' If line_info(0) = "PREV POST LIST" 							Then ObjExcel.Cells(excel_row, prev_post_list_col_const).Value  = line_info(1)
				' If line_info(0) = "PREV POST CURR VERIF POST" 				Then ObjExcel.Cells(excel_row, prev_post_curr_verif_post_col_const).Value  = line_info(1)
				' If line_info(0) = "PREV POST ONGOING SNAP APP" 				Then ObjExcel.Cells(excel_row, prev_post_reg_snap_app_col_const).Value  = line_info(1)
				' If line_info(0) = "PREV POST VERIFS RECVD" 					Then ObjExcel.Cells(excel_row, prev_post_verifs_recvd_col_const).Value  = line_info(1)
				' If line_info(0) = "EXPLAIN APPROVAL DELAYS"                 Then ObjExcel.Cells(excel_row, expl_appr_delay_col_const).Value = line_info(1)
				' If line_info(0) = "POSTPONED VERIFICATIONS"                 Then ObjExcel.Cells(excel_row, post_verifs_yn_col_const).Value = line_info(1)
				' If line_info(0) = "WHAT ARE THE POSTPONED VERIFICATIONS"    Then ObjExcel.Cells(excel_row, post_verifs_list_col_const).Value = line_info(1)
				' If line_info(0) = "FACI DELAY ACTION" 						Then ObjExcel.Cells(excel_row, faci_delay_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI DENY" 								Then ObjExcel.Cells(excel_row, faci_deny_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI NAME" 								Then ObjExcel.Cells(excel_row, faci_name_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI INELIG SNAP" 						Then ObjExcel.Cells(excel_row, faci_snap_inelig_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI ENTRY DATE" 						Then ObjExcel.Cells(excel_row, faci_entry_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI RELEASE DATE" 						Then ObjExcel.Cells(excel_row, faci_release_col_const).Value  = line_info(1)
				' If line_info(0) = "FACI RELEASE IN 30 DAYS" 				Then ObjExcel.Cells(excel_row, faci_release_in_30_col_const).Value  = line_info(1)
				' If line_info(0) = "DATE OF SCRIPT RUN"                      Then ObjExcel.Cells(excel_row, script_run_date_col_const).Value = line_info(1)
				' If line_info(0) = "SCRIPT RUN"                      		Then ObjExcel.Cells(excel_row, script_run_col_const).Value = line_info(1)

				'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW
				' If line_info(0) = "CASE NUMBER"                             Then ObjDISCExcel.Cells(excel_row, disc_case_nbr_col_const).Value  = line_info(1)
				' If line_info(0) = "EXPEDITED DETERMINATION STATUS"          Then ObjDISCExcel.Cells(excel_row, disc_exp_det_col_const).Value  = line_info(1)
				' If line_info(0) = "DATE OF APPLICATION" Then
				' 	ObjDISCExcel.Cells(excel_row, disc_appl_date_col_const).Value  = line_info(1)
				' 	appl_date = line_info(1)
				' End If
				'
				' If line_info(0) = "APPT NOTC SENT DATE" Then
				' 	ObjDISCExcel.Cells(excel_row, disc_notc_date_col_const).Value  = line_info(1)
				' 	notc_date = line_info(1)
				' End If
				'
				' If line_info(0) = "DATE OF INTERVIEW" Then
				' 	ObjDISCExcel.Cells(excel_row, disc_intv_date_col_const).Value  = line_info(1)
				' 	intvw_date = line_info(1)
				' End If
				' If line_info(0) = "DATE OF APPROVAL" Then
				' 	ObjDISCExcel.Cells(excel_row, disc_app_date_col_const).Value  = line_info(1)
				' 	app_date = line_info(1)
				' End If
				' If line_info(0) = "ID ON FILE"                              Then ObjDISCExcel.Cells(excel_row, disc_id_col_const).Value = line_info(1)
				' If line_info(0) = "EXPLAIN APPROVAL DELAYS"                 Then ObjDISCExcel.Cells(excel_row, disc_app_delays_col_const).Value = line_info(1)
				' If line_info(0) = "DET INCOME" 								Then ObjDISCExcel.Cells(excel_row, disc_income_col_const).Value  = line_info(1)
				' If line_info(0) = "DET ASSETS" 								Then ObjDISCExcel.Cells(excel_row, disc_asset_col_const).Value  = line_info(1)
				' If line_info(0) = "DET SHEL" 								Then ObjDISCExcel.Cells(excel_row, disc_shelter_col_const).Value  = line_info(1)
				' If line_info(0) = "DET HEST" 								Then ObjDISCExcel.Cells(excel_row, disc_utilities_col_const).Value  = line_info(1)
				' If line_info(0) = "EXPEDITED SCREENING STATUS"              Then ObjDISCExcel.Cells(excel_row, disc_screening_col_const).Value  = line_info(1)
				' If line_info(0) = "WORKER NAME"                             Then ObjDISCExcel.Cells(excel_row, disc_worker_col_const).Value  = line_info(1)
				' If line_info(0) = "SCRIPT RUN"                      		Then ObjDISCExcel.Cells(excel_row, disc_script_run_col_const).Value = line_info(1)

				If save_to_worklist = True Then
					If line_info(0) = "CASE NUMBER" 					Then ObjWORKExcel.Cells(work_excel_row, work_case_nbr_col_const).Value = line_info(1)
					If line_info(0) = "WORKER NAME" 					Then ObjWORKExcel.Cells(work_excel_row, work_worker_col_const).Value = line_info(1)
					If line_info(0) = "DATE OF APPLICATION" 			Then ObjWORKExcel.Cells(work_excel_row, work_appl_date_col_const).Value = line_info(1)
					If line_info(0) = "APPT NOTC SENT DATE" 			Then ObjWORKExcel.Cells(work_excel_row, work_notc_date_col_const).Value = line_info(1)
					If line_info(0) = "DATE OF INTERVIEW" 				Then ObjWORKExcel.Cells(work_excel_row, work_intv_date_col_const).Value = line_info(1)
					If line_info(0) = "DATE OF APPROVAL" 				Then ObjWORKExcel.Cells(work_excel_row, work_app_date_col_const).Value = line_info(1)
					If line_info(0) = "ID ON FILE" 						Then ObjWORKExcel.Cells(work_excel_row, work_id_col_const).Value = line_info(1)
					If line_info(0) = "EXPLAIN APPROVAL DELAYS" 		Then ObjWORKExcel.Cells(work_excel_row, work_app_delays_col_const).Value = line_info(1)
					If line_info(0) = "EXPEDITED DETERMINATION STATUS" 	Then ObjWORKExcel.Cells(work_excel_row, work_exp_det_col_const).Value = line_info(1)
					If line_info(0) = "DET INCOME" 						Then ObjWORKExcel.Cells(work_excel_row, work_income_col_const).Value = line_info(1)
					If line_info(0) = "DET ASSETS" 						Then ObjWORKExcel.Cells(work_excel_row, work_asset_col_const).Value = line_info(1)
					If line_info(0) = "DET SHEL" 						Then ObjWORKExcel.Cells(work_excel_row, work_shelter_col_const).Value = line_info(1)
					If line_info(0) = "DET HEST" 						Then ObjWORKExcel.Cells(work_excel_row, work_utilities_col_const).Value = line_info(1)
					If line_info(0) = "DATE OF SCRIPT RUN"              Then ObjWORKExcel.Cells(work_excel_row, work_script_run_date_col_const).Value = line_info(1)

				End If
			End If
		Next

		' If create_a_test_worklist = False Then
		'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW
		' If IsDate(appl_date) = True Then appl_date = DateAdd("d", 0, appl_date)
		' If IsDate(notc_date) = True Then notc_date = DateAdd("d", 0, notc_date)
		' If IsDate(intvw_date) = True Then intvw_date = DateAdd("d", 0, intvw_date)
		' If IsDate(app_date) = True Then app_date = DateAdd("d", 0, app_date)
		'
		' If IsDate(appl_date) = True AND IsDate(notc_date) = True Then
		' 	count_days = 0
		' 	Do While DateDiff("d", appl_date, notc_date) <> 0
		' 		appl_date = DateAdd("d", 1, appl_date)
		' 		call change_date_to_soonest_working_day(appl_date, "FORWARD")
		' 		If DateDiff("d", appl_date, notc_date) < 0 Then Exit Do
		' 		count_days = count_days + 1
		' 	Loop
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_appl_notc_col_const).Value = count_days
		' Else
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_appl_notc_col_const).Value = ""
		' End If
		'
		' If IsDate(notc_date) = True AND IsDate(intvw_date) = True Then
		' 	count_days = 0
		' 	If notc_date <= intvw_date Then
		' 		Do While DateDiff("d", notc_date, intvw_date) <> 0
		' 			notc_date = DateAdd("d", 1, notc_date)
		' 			call change_date_to_soonest_working_day(notc_date, "FORWARD")
		' 			If DateDiff("d", notc_date, intvw_date) < 0 Then Exit Do
		' 			count_days = count_days + 1
		' 		Loop
		' 	End If
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_notc_intv_col_const).Value = count_days
		' Else
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_notc_intv_col_const).Value = ""
		' End If
		'
		' If IsDate(appl_date) = True AND IsDate(app_date) = True Then
		' 	count_days = 0
		' 	Do While DateDiff("d", appl_date, app_date) <> 0
		' 		appl_date = DateAdd("d", 1, appl_date)
		' 		call change_date_to_soonest_working_day(appl_date, "FORWARD")
		' 		If DateDiff("d", appl_date, app_date) < 0 Then Exit Do
		' 		count_days = count_days + 1
		' 	Loop
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_appl_app_col_const).Value = count_days
		' Else
		' 	ObjDISCExcel.Cells(excel_row, disc_bdays_appl_app_col_const).Value = ""
		' End If
		' End If
		excel_row = excel_row + 1													'advance to the next row

		STATS_counter = STATS_counter + 1
		If save_to_worklist = True Then work_excel_row = work_excel_row + 1
		objTextStream.Close						'we are done with this file, so we must close the access
	    objFSO.DeleteFile(this_file_path)		'deleting the TXT file because hgave the information
	Next

	If create_a_test_worklist = True then
		objWORKWorkbook.Save()		'saving the excel

		If leave_excel_open = "No - Close the file" Then		'if the file should be closed - it does it here.
			ObjWORKExcel.ActiveWorkbook.Close

			ObjWORKExcel.Application.Quit
			ObjWORKExcel.Quit
		Else
			ObjWORKExcel.Visible = True
		End If
	End If

	' REmoving the DAY specific sheet in the Master List'
	' 'formatting the worksheet made for the day
	' objTODAYWorkSheet.Activate
	' Const xlSrcRange = 1
	' Const xlYes = 1
	' xlVAlignTop = -4160
	' xlHAlignLeft = -4131
	' For col = 1 to 17
	'     ObjExcel.columns(col).AutoFit()
	'     ObjExcel.columns(col).VerticalAlignment = xlVAlignTop
	'     ObjExcel.columns(col).HorizontalAlignment = xlHAlignLeft
	' Next
	'
	' 'setting some column widths on the day sheet
	' ObjExcel.Columns(deny_reason_col_const).ColumnWidth = 150
	' ObjExcel.Columns(deny_reason_col_const).WrapText = True
	' ObjExcel.Columns(expl_appr_delay_col_const).ColumnWidth = 150
	' ObjExcel.Columns(expl_appr_delay_col_const).WrapText = True
	'
	' 'here we add the table format to the sheet from today
	' tableRange = "A1:AN" & excel_row-1
	' table_friendly_date = replace(date, "/", "")
	' table_friendly_date = trim(table_friendly_date)
	' table_name = table_friendly_date & "TABLE"
	' ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, tableRange, xlYes).Name = table_name
	' ObjExcel.ActiveSheet.ListObjects(table_name).TableStyle = "TableStyleDark2"

	'ending on the statistics sheet
	For Each objWorkSheet In objWorkbook.Worksheets
	    If objWorkSheet.Name = "Statistics" Then
			objWorkSheet.Activate
	        Exit For
	    End If
	Next

	objWorkbook.Save()		'saving the excel
	' objDISCWorkbook.Save()		'saving the excel		'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW

	objExcel.ActiveWorkbook.Close
	' ObjDISCExcel.ActiveWorkbook.Close					'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW

	objExcel.Application.Quit
	objExcel.Quit
	' ObjDISCExcel.Application.Quit						'THIS IS FOR THE FILE - JAKE'S DISCOVERY WHICH IS NOT NEEDED RIGHT NOW
	' ObjDISCExcel.Quit
End If

If report_selection = "Combine Worklists" Then

	last_week = DateAdd("d", -7, date)
	day_of_week = weekday(last_week)
	adjust_to_sunday = 1 - day_of_week
	adjust_to_saturday = 7 - day_of_week

	last_week_sunday = DateAdd("d", adjust_to_sunday, last_week)
	last_week_saturday = DateAdd("d", adjust_to_saturday, last_week)
	' MsgBox "Last Week: Sunday - " & last_week_sunday & vbCr & "                   Saturday - " & last_week_saturday

	'Open the worklist report
	Call excel_open(worklist_review_file, True, True, ObjReportExcel, objReportWorkbook)  			'opens the selected excel file'

	For Each objWorkSheet In objReportWorkbook.Worksheets									'looking through each of the worksheets to find the 'ALL CASES' worksheet
		If instr(objWorkSheet.Name, "All Review Cases") <> 0 Then
			set objALLCASESWorkSheet = objWorkSheet									'setting the 'ALL CASES' to a worksheet variable because we need it a lot
			objALLCASESWorkSheet.Activate											'opening that worksheet
			Exit For
		End If
	Next

	'Now we need to find the last row in the 'ALL CASES' sheet so we don't overwrite anything
	total_excel_row = 1																'default to the first row
	Do
		total_excel_row = total_excel_row + 1
		this_case_number = trim(ObjReportExcel.Cells(total_excel_row, 1).Value)
	Loop Until this_case_number = ""												'if the case number is blank then the row is blank

	'create a sheet for last week'

	'Add a sheet to the Excel with the report date
	sheet_name = "Week of " & last_week_sunday
	ObjReportExcel.Worksheets.Add().Name = sheet_name

	For Each objWorkSheet In objReportWorkbook.Worksheets									'setting the worksheet to a variable so we can use it again
	    If objWorkSheet.Name = sheet_name Then
			set objTODAYWorkSheet = objWorkSheet
			objTODAYWorkSheet.Activate
	        Exit For
	    End If
	Next
	excel_row = 2

	'Add Column Headers
	'



	Set objWorkFolder = objFSO.GetFolder(worklist_template_folder)										'Creates an oject of the whole my documents folder
	Set colWorkFiles = objWorkFolder.Files																'Creates an array/collection of all the files in the folder
	For Each objWorkFile in colWorkFiles																'looping through each file
		this_file_name = objWorkFile.Name															'Grabing the file name
		this_file_type = objWorkFile.Type															'Grabing the file type
		this_file_created_date = objWorkFile.DateCreated											'Reading the date created
		this_file_path = objWorkFile.Path															'Grabing the path for the file

		worklist_from_last_week = False
		If this_file_created_date >= last_week_sunday and this_file_created_date <= last_week_saturday Then worklist_from_last_week = True

		If worklist_from_last_week = True Then
			MsgBox "File name - " & this_file_name
			'Open the Excel file - not visible
			'copy each line into the All Cases sheet
			'copy each line into the sheet for the week
			'close the file

			' objFSO.MoveFile this_file_path , worklist_archive_path & "\" & this_file_name & ".xlsx"    'moving each file to the archive file

		End If

	Next

	'turn the new sheet to a table
	'save all files
	MsgBox "Stop here"
End If

'SAVE EXCEL'
Call script_end_procedure("Expedited Determination report is updated and the tracking files removed.")
