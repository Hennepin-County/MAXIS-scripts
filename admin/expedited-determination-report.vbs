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
const worker_user_id_col_const				= 3
const xnumber_col_const 					= 4
const date_of_appl_col_const 				= 5
const appt_notc_date_col_const				= 6
const date_of_appt_col_const				= 7
const date_of_intve_col_const 				= 8
const screen_status_col_const 				= 9
const det_status_col_const 					= 10
const det_income_col_const 					= 11
const det_asset_col_const 					= 12
const det_shel_col_const 					= 13
const det_hest_col_const 					= 14
const date_of_appr_col_const 				= 15
const date_of_deny_col_const 				= 16
const deny_reason_col_const 				= 17
const id_on_file_col_const 					= 18
const outstate_action_col_const 			= 19
const outstate_state_col_const 				= 20
const outstate_end_date_rept_col_const 		= 21
const outstate_openended_col_const 			= 22
const outstate_end_date_verif_col_const 	= 23
const mn_elig_begin_col_const 				= 24
const prev_post_delay_col_const 			= 25
const prev_post_prev_date_of_appl_col_const = 26
const prev_post_list_col_const 				= 27
const prev_post_curr_verif_post_col_const 	= 28
const prev_post_reg_snap_app_col_const 		= 29
const prev_post_verifs_recvd_col_const 		= 30
const expl_appr_delay_col_const 			= 31
const post_verifs_yn_col_const 				= 32
const post_verifs_list_col_const 			= 33
const faci_delay_col_const 					= 34
const faci_deny_col_const 					= 35
const faci_name_col_const 					= 36
const faci_snap_inelig_col_const 			= 37
const faci_entry_col_const 					= 38
const faci_release_col_const 				= 39
const faci_release_in_30_col_const 			= 40
const script_run_date_col_const 			= 41
const script_run_col_const					= 42

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
const exch_rept_id_col_const 				= 15 	' ID
const exch_rept_faci_col_const 				= 16 	' FACI
const exch_rept_out_of_state_col_const 		= 17 	' Out of State
const exch_rept_prev_exp_col_const 			= 18 	' Previous EXP Not Verified
const exch_rept_deu_disq_col_const 			= 19 	' DEU/DISQ
const exch_rept_imig_col_const 				= 20 	' Immigration
const exch_rept_new_hire_col_const 			= 21 	' New Hire
const exch_rept_job_info_col_const 			= 22 	' Attested Income/STWK
const exch_rept_other_info_col_const 		= 23 	' Other Attested Information
const exch_rept_insuf_intvw_col_const 		= 24 	' Did not gather enough informat at interview
const exch_rept_HSR_lacks_support_col_const = 25 	' Worker Lacks Support
const exch_rept_insuf_case_note_col_const 	= 26 	' Was CASE/NOTE Sufficient
const exch_rept_MAXIS_updated_col_const 	= 27 	' Was MAXIS Updated Correctly
const exch_rept_HSR_knew_poli_col_const 	= 28 	' Worker Knew Policy
const exch_rept_exchange_col_const 			= 29 	' Exchange Needed?
const exch_rept_exch_date_time_col_const 	= 30 	' Date/Time of Exchange - When did you connect with the worker
const exch_rept_exch_durr_col_const 		= 31 	' Durration of Exchange
const exch_rept_unable_to_connect_col_const = 32 	' Unable to Connect
const exch_rept_notes_col_const 			= 33 	' Notes
const exch_worklist_date_time_col_const		= 34
const exch_app_date_time_col				= 35
const exch_app_status_col 					= 36
const exch_app_exp_status_col 				= 37


'END DECLARATIONS BLOCK ====================================================================================================

EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
Call check_for_MAXIS(true)
leave_excel_open = "No - Close the file"

'Declaring the only dialog
Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 246, 100, "Expedited Determination Report"
	  DropListBox 10, 80, 180, 45, "Yes - Keep the file open"+chr(9)+"No - Close the file", leave_excel_open
	  ButtonGroup ButtonPressed
	    OkButton 200, 80, 40, 15
	  Text 10, 10, 225, 30, "This script is used to pull reports around information gathered during the Expedited Determination script runs to provide insight in how we are handling Expedited SNAP in Hennepin County"
	  Text 10, 45, 225, 10, "When the script is complete, the Excel will be saved."
	  Text 10, 60, 130, 20, "At the end of the script run, would you like the Excel file to remain open:"
	EndDialog

	'showing the dialog - there is no loop because there is nothing to manage and no password handling.
	dialog Dialog1
	cancel_confirmation

	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = False					'loops until user passwords back in

'defining the assignment folder and setting the file paths
exp_assignment_folder = t_drive & "\Eligibility Support\Assignments\Expedited Information"
Set objFolder = objFSO.GetFolder(exp_assignment_folder)										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder

txt_file_archive_path = t_drive & "\Eligibility Support\Assignments\Expedited Information\Archive"

report_out_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP Determination Report Out.xlsx"
hss_report_file = "C:\Users\" & user_ID_for_validation & "\Hennepin County\ES Management - Documents\Case Review\EXP Determination HSS Report.xlsx"

'HERE we REVIEW ALL THE TXT FILES
	'If EXPEDITED - check MAXIS to see if SNAP is still pending or not.

Call excel_open(report_out_file, True, True, ObjExcel, objWorkbook)  			'opens the selected excel file'
Call excel_open(hss_report_file, True, True, ObjHSSExcel, objHSSWorkbook)  			'opens the selected excel file'

const hss_rept_report_day_col			= 1
const hss_rept_case_number_col			= 2
const hss_rept_application_date_col		= 3
const hss_rept_interview_date_col		= 4
const hss_rept_approval_delay_detail_col= 5
const hss_rept_script_user_col			= 6
const hss_rept_hsr_email_col			= 7
const hss_rept_script_user_id_col		= 8
const hss_rept_hss_name_col				= 9
const hss_rept_hss_email_col			= 10
const hss_rept_pm_name_col				= 11
const hss_rept_pm_email_col				= 12
const hss_rept_total_report_row_col		= 13

const worker_name_col 	= 1
const worker_mx_id_col 	= 2
const worker_hc_id_col 	= 3
const worker_email_col 	= 4
const hss_name_col 		= 5
const hss_email_col 	= 6
const pm_name_col 		= 7
const pm_email_col 		= 8

const hsr_name_const 	= 0
const hsr_mx_id_const 	= 1
const hsr_hc_id_const	= 2
const hsr_email_const 	= 3
const hss_name_const 	= 4
const hss_email_const 	= 5
const pm_name_const 	= 6
const pm_email_const 	= 7

'This is where we get the information about HSRs and HSSs - we need to determine the data source and update this functionality once received - currently it is using a sheet in the HSS report out Excel File
Dim WORKER_ARRAY()
ReDim WORKER_ARRAY(pm_email_const, 0)

ObjHSSExcel.worksheets("HSS List").Activate

excel_row = 2
worker_count = 0
Do
	ReDim preserve WORKER_ARRAY(pm_email_const, worker_count)

	WORKER_ARRAY(hsr_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, worker_name_col))
	WORKER_ARRAY(hsr_mx_id_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, worker_mx_id_col))
	WORKER_ARRAY(hsr_hc_id_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, worker_hc_id_col))
	WORKER_ARRAY(hsr_email_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, worker_email_col))
	WORKER_ARRAY(hss_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, hss_name_col))
	WORKER_ARRAY(hss_email_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, hss_email_col))
	WORKER_ARRAY(pm_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, pm_name_col))
	WORKER_ARRAY(pm_email_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, pm_email_col))

	excel_row = excel_row + 1
	worker_count = worker_count + 1
	next_worker_info = trim(ObjHSSExcel.Cells(excel_row, worker_name_col))
Loop until next_worker_info = ""

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

'Now we are going to read all of the txt files generated from the Expedited Determination script runs. If they need review, they will be added to the report out file - otherwise they will be saved in the large Excel list
ObjHSSExcel.worksheets("Case Email List").Activate							'activating the sheet

'finding the last row of the HSS Report Out ffile
hss_excel_row = 1																'default to the first row
Do
	hss_excel_row = hss_excel_row + 1											'increment to the next row
	this_case_number = trim(ObjHSSExcel.Cells(hss_excel_row, 2).Value)			'check to see if there is information on this row - if not, the script will leave the loop and know the first blank row
	row_deleted = False															'defaulting the row to NOT being deleted.

	'deleting old rows - historical information is not saved in this excel, just pulled from it and added to a sharepoint list.
	report_day = trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_report_day_col).Value)		'this is the day the row was created.
	If report_day <> "" Then													'if the report day is not blank we are going to check and see if it is old.
		report_day = DateAdd("d", 0, report_day)								'first we make sure it is a date
		If DateDiff("d", report_day, date) >= 30 Then							'if that date was more than 30 days ago, we don't need it.
			row_deleted = True													'if more than 30 days old, we are deleting the row
			SET objRange = ObjHSSExcel.Cells(hss_excel_row, 1).EntireRow		'select the row
			objRange.Delete														'delete the selection
			hss_excel_row = hss_excel_row - 1									'now we go back a row, since this one no longer exists.
		End If
	End If
	'here we are going to fill in the HSR information if it is missing
	If row_deleted = False Then													'if the row was deleted, there is nothing else to do with it
		If trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_hsr_email_col).Value) = "" Then	'if the email is blank, add it from the worker array'
			For each_wrkr = 0 to UBound(WORKER_ARRAY, 2)			'here we need to use the data of HSRs and HSSs to fill in the appropriate HSS and PM based on Worker Name
				If WORKER_ARRAY(hsr_hc_id_const, each_wrkr) = trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_id_col).Value) Then
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_col).Value = WORKER_ARRAY(hsr_name_const, each_wrkr)
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_hsr_email_col).Value = WORKER_ARRAY(hsr_email_const, each_wrkr)
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_name_col).Value = WORKER_ARRAY(hss_name_const, each_wrkr)
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_email_col).Value = WORKER_ARRAY(hss_email_const, each_wrkr)
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_name_col).Value = WORKER_ARRAY(pm_name_const, each_wrkr)
					ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_email_col).Value = WORKER_ARRAY(pm_email_const, each_wrkr)
				End If
			Next
		End If
	End If
	report_day = ""
Loop Until this_case_number = ""

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

missing_HSRs = ""
'Looking at each txt file in the assignments folder to capture Expedited Determination information
For Each objFile in colFiles																'looping through each file
	report_to_HSS = False														'setting some default variables
	exp_det = False
	case_nbr_hold = ""
    this_file_path = objFile.Path												'identifying the current file
	this_file_name = objFile.Name
	this_file_created_date = objFile.DateCreated								'Reading the date created

	If DateDiff("d", this_file_created_date, date) > 0 Then						'we are only pulling information for cases that were reviewed yesterday at this time.
	    'Setting the object to open the text file for reading the data already in the file
	    Set objTextStream = objFSO.OpenTextFile(this_file_path, ForReading)

	    'Reading the entire text file into a string
	    every_line_in_text_file = objTextStream.ReadAll

	    exp_det_details = split(every_line_in_text_file, vbNewLine)					'creating an array of all of the information in the TXT files

		For Each text_line in exp_det_details										'read each line in the file
			If Instr(text_line, "^*^*^") <> 0 Then
				line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
				line_info(0) = trim(line_info(0))
				'here we add the information from TXT to Excel

				If line_info(0) = "CASE NUMBER"                             Then ObjExcel.Cells(total_excel_row, case_number_col_const).Value  = line_info(1)
				If line_info(0) = "WORKER NAME"                             Then ObjExcel.Cells(total_excel_row, worker_col_const).Value  = line_info(1)
				If line_info(0) = "WORKER USER ID"                          Then ObjExcel.Cells(total_excel_row, worker_user_id_col_const).Value  = line_info(1)
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
					If UCASE(line_info(1))&"" = "TRUE" Then exp_det = True		'identifying if the case appeared expedited at the time of the Expedited Determination script run
				End If
				If line_info(0) = "CASE NUMBER" Then case_nbr_hold = line_info(1)	''saving the case number so we can do things in a second
			End If
		Next
		'if the case is expedited, we need to see if it is pending
		If exp_det = True Then
			MAXIS_case_number = case_nbr_hold									'setting the case number
			Call back_to_SELF
			Call navigate_to_MAXIS_screen("CASE", "CURR")						'go check CASE CURR - we are looking for FS PENDING
			row = 1
			col = 1
			EMSearch "FS:", row, col
			If row <> 0 Then									'If we found FS listed - we will look for the current program status
				EMReadScreen fs_status, 9, row, col + 4
				fs_status = trim(fs_status)
				If fs_status = "PENDING" Then report_to_HSS = True			'if the program status is PENDING, then action has not been taken and the case needs to be reviewed.
			End If
			Call back_to_SELF
			MAXIS_case_number = ""			''blanking out the case number
		End If

		''if determined that the case needs to be reviewed - then we need to add it to the Excel for the HSS Report Out
		If report_to_HSS = True Then
			For Each text_line in exp_det_details										'read each line in the file
				If Instr(text_line, "^*^*^") <> 0 Then
					line_info = split(text_line, "^*^*^")								'creating a small array for each line. 0 has the header and 1 has the information
					line_info(0) = trim(line_info(0))

					'we only save relevant information to the HSS Report Out Excel
					If line_info(0) = "CASE NUMBER" 			Then ObjHSSExcel.Cells(hss_excel_row, hss_rept_case_number_col).Value = line_info(1)
					If line_info(0) = "DATE OF APPLICATION" 	Then ObjHSSExcel.Cells(hss_excel_row, hss_rept_application_date_col).Value  = line_info(1)
					If line_info(0) = "DATE OF INTERVIEW" 		Then ObjHSSExcel.Cells(hss_excel_row, hss_rept_interview_date_col).Value = line_info(1)
					If line_info(0) = "EXPLAIN APPROVAL DELAYS" Then ObjHSSExcel.Cells(hss_excel_row, hss_rept_approval_delay_detail_col).Value = line_info(1)
					If line_info(0) = "WORKER USER ID" Then
						ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_id_col).Value = line_info(1)
						For each_wrkr = 0 to UBound(WORKER_ARRAY, 2)			'here we need to use the data of HSRs and HSSs to fill in the appropriate HSS and PM based on Worker Name
							If WORKER_ARRAY(hsr_hc_id_const, each_wrkr) = line_info(1) Then
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_col).Value = WORKER_ARRAY(hsr_name_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_hsr_email_col).Value = WORKER_ARRAY(hsr_email_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_name_col).Value = WORKER_ARRAY(hss_name_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_email_col).Value = WORKER_ARRAY(hss_email_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_name_col).Value = WORKER_ARRAY(pm_name_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_email_col).Value = WORKER_ARRAY(pm_email_const, each_wrkr)
							End If
						Next
						If ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_col).Value = "" Then missing_HSRs = missing_HSRs & vbCr & line_info(1)
					End If

				End If
			Next
			ObjHSSExcel.Cells(hss_excel_row, hss_rept_total_report_row_col).Value = total_excel_row
			ObjHSSExcel.Cells(hss_excel_row, hss_rept_report_day_col).Value = date
			hss_excel_row = hss_excel_row + 1
		End If

		total_excel_row = total_excel_row + 1										'advance to the next row
		STATS_counter = STATS_counter + 1

		objTextStream.Close						'we are done with this file, so we must close the access

		Dim oTxtFile
		With (CreateObject("Scripting.FileSystemObject"))
			'If the file exists in the archive, we we will delete the version in the archive so the one from the main file can be placed in archive
			If .FileExists(txt_file_archive_path & "\" & this_file_name & ".txt") Then
				objFSO.DeleteFile(txt_file_archive_path & "\" & this_file_name & ".txt")		'deleting the TXT file because hgave the information
			End If
		End With
		' On error resume next
		objFSO.MoveFile this_file_path , txt_file_archive_path & "\" & this_file_name & ".txt"    'moving each file to the archive file
		' If Err.Number <> 0 Then MsgBox "FILE IS DUPLICATE ???" & vbCr & "this_file_path - " & this_file_path & vbCr & "archive pather - " & txt_file_archive_path & "\" & this_file_name & ".txt"
		' On Error Goto 0
	End If
Next
objWorkbook.Save()		'saving the excel
objHSSWorkbook.Save()		'saving the excel

'closing all the files if requested'
If leave_excel_open = "No - Close the file" Then
	ObjHSSExcel.ActiveWorkbook.Close

	ObjHSSExcel.Application.Quit
	ObjHSSExcel.Quit

	objExcel.ActiveWorkbook.Close

	objExcel.Application.Quit
	objExcel.Quit
End If

'This will send an email if HSR information is missing in the data
If missing_HSRs <> "" Then
	email_subject = "Missing HSR Detial in HSS Expedited Report"

	email_body = "The Expedited Determination Report script has run and was unable to match the HSR User ID in at least one instance. Review the Worker List sheet in the Report Excel and update to include the followin User ID(s):"
	email_body = email_body & missing_HSRs
	email_body = email_body & vbCr & vbCr & "This email is automated as a part of the script run of ADMIN - Expedited Determination Report."

	send_email = True
	Call create_outlook_email("", "HSPH.EWS.BlueZoneScripts@hennepin.us", "", "", email_subject, 1, False, "", "", False, "", email_body, False, "", send_email)
End If

' Here we go and delete the txt files that are generated with the Exp Det script run from the archive files IF the file is more than 2 weeks old.
last_week = DateAdd("d", -14, date)
day_of_week = weekday(last_week)
adjust_to_sunday = 1 - day_of_week
adjust_to_saturday = 7 - day_of_week

last_week_saturday = DateAdd("d", adjust_to_saturday, last_week)

'ONCE THIS IS ALL DONE ADD FUNCTIONALITY TO DELETE ALL THE FILES IN THE ARCHIVE FOLDER OLDER THAN THE CURRENT WEEK - Since we know those are recorded.'
Set objTXTArchiveFolder = objFSO.GetFolder(txt_file_archive_path)										'Creates an oject of the whole my documents folder
Set colTXTArchiveFiles = objTXTArchiveFolder.Files																'Creates an array/collection of all the files in the folder

For Each objFile in colTXTArchiveFiles																'looping through each file
	this_file_path = objFile.Path
	this_file_name = objFile.Name
	this_file_created_date = objFile.DateCreated											'Reading the date created

	If DateDiff("d", this_file_created_date, last_week_saturday) >=0 Then
		objFSO.DeleteFile(this_file_path)		'deleting the TXT file because hgave the information
	End If
Next

'SAVE EXCEL'
Call script_end_procedure("Expedited Determination report is updated and the tracking files removed.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------07/14/2022
'--Tab orders reviewed & confirmed----------------------------------------------07/14/2022
'--Mandatory fields all present & Reviewed--------------------------------------N/A
'--All variables in dialog match mandatory fields-------------------------------N//A
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------07/14/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------07/14/2022
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------07/14/2022
'--BULK - review output of statistics and run time/count (if applicable)--------07/14/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---07/14/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------07/14/2022
'--Incrementors reviewed (if necessary)-----------------------------------------07/14/2022
'--Denomination reviewed -------------------------------------------------------07/14/2022
'--Script name reviewed---------------------------------------------------------07/14/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------07/14/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------07/14/2022
'--comment Code-----------------------------------------------------------------07/14/2022
'--Update Changelog for release/update------------------------------------------07/14/2022
'--Remove testing message boxes-------------------------------------------------07/14/2022
'--Remove testing code/unnecessary code-----------------------------------------07/14/2022
'--Review/update SharePoint instructions----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------07/14/2022
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------N/A
