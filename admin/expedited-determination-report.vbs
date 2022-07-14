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
' discovery_template_worklist_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Jake's Discovery\"
' worklist_archive_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists\Archive\"
' worklist_template_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists"
'
' discovery_template_file = discovery_template_worklist_path & "Discovery Template.xlsx"
' worklist_template_file = worklist_template_path & "Worklist Template.xlsx"

report_out_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP Determination Report Out.xlsx"
worklist_template_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Exp Det Worklists\"
worklist_review_file = worklist_template_path & "Worklist Review Report.xlsx"
hss_report_file = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP Determination HSS Report.xlsx"

'OPEN the Worklist Report and check for approvals.
'Open the worklist report
Call excel_open(worklist_review_file, True, True, ObjReportExcel, objReportWorkbook)  			'opens the selected excel file'

For Each objWorkSheet In objReportWorkbook.Worksheets									'looking through each of the worksheets to find the 'ALL CASES' worksheet
	If instr(objWorkSheet.Name, "All Review Cases") <> 0 Then
		set objALLCASESWorkSheet = objWorkSheet									'setting the 'ALL CASES' to a worksheet variable because we need it a lot
		objALLCASESWorkSheet.Activate											'opening that worksheet
		Exit For
	End If
Next

'In this part we look at the Excel file from the Expedited Exchange project worklist to find the approval information
total_excel_row = 2
Do
	version_number = ""			''resetting variables
	process_date = ""
	elig_result = ""
	approval_status = ""
	note_time = ""
	If trim(ObjReportExcel.Cells(total_excel_row, exch_app_date_time_col).Value) = "" Then				''Only look at cases where the approval has not been found

		MAXIS_case_number = trim(ObjReportExcel.Cells(total_excel_row, work_case_nbr_col_const).Value)			'grabbing the CASE Number'
		date_of_application = ObjReportExcel.Cells(total_excel_row, work_appl_date_col_const).Value				'getting the date of application
		Call convert_date_into_MAXIS_footer_month(date_of_application, MAXIS_footer_month, MAXIS_footer_year)	'use the app date to get the footer month and year'
		date_of_application = DateAdd("d", 0, date_of_application)				'making sure this is a date

		Call back_to_SELF
		Call navigate_to_MAXIS_screen("ELIG", "FS  ")							'go check ELIG FS and find an approved version
		EMWriteScreen "99", 19, 78
		transmit

		approved_version_found = False
		elig_row = 17
		Do
			EMReadScreen version_number, 2, elig_row, 22
			EMReadScreen process_date, 8, elig_row, 26
			EMReadScreen elig_result, 11, elig_row, 37
			EMReadScreen approval_status, 10, elig_row, 50
			' MsgBox process_date
			version_number = trim(version_number)
			elig_result = trim(elig_result)
			approval_status = trim(approval_status)
			If approval_status = "APPROVED" Then
				process_date = DateAdd("d", 0, process_date)
				' MsgBox "date_of_application - " & date_of_application & vbCr & "process_date - " & process_date & vbCr & "date diff - " & DateDiff("d", date_of_application, process_date)
				If DateDiff("d", date_of_application, process_date) >=0 Then
					approved_version_found = True

					version_number = version_number & "  "
					EMWriteScreen version_number, 18, 54
					' MsgBox "Pause"
					transmit

					EMReadScreen auto_close_warning, 11, 11, 43
					If auto_close_warning = "Auto-Closed" Then
						approved_version_found = False
						transmit
						Call navigate_to_MAXIS_screen("ELIG", "FS  ")
						EMWriteScreen "99", 19, 78
						transmit
					Else
						exit Do
					End If
				End If
			End If

			elig_row = elig_row - 1
		Loop until elig_row = 6

		If approved_version_found = True Then									'if an approval was found, we need to capture some details
			EMReadScreen approved_date, 8, 3, 14								'date approved
			' MsgBox approved_date
			approved_date = DateAdd("d", 0, approved_date)

			Call write_value_and_transmit("FSCR", 19, 70)						'determining if the approval was expedited
			EMReadScreen expedited_status, 9, 4, 3
			expedited_status = trim(expedited_status)

			Call write_value_and_transmit("FSSM", 19, 70)						'determine if the approval made the case ELIGIBLE or INELIGIBLE
			EMReadScreen elig_status, 10, 7, 31
			elig_status = trim(elig_status)

			Call navigate_to_MAXIS_screen("CASE", "NOTE")						'go to CNOTE to see if we can find the approved TIME
			too_old_date = DateAdd("d", -1, approved_date)						'we don''t want to look past the date FS was approved
			note_row = 5
			Do
				EMReadScreen note_date, 8, note_row, 6                  		'reading the note date
				EMReadScreen part_note_title, 11, note_row, 25               	'reading the note header
				EMReadScreen full_note_title, 55, note_row, 25               	'reading the note header
				note_date = DateAdd("d", 0, note_date)

				If DateDiff("d", note_date, approved_date) = 0 Then				'If the CNOTE was created the date approved
					If (part_note_title = "---Approved" or part_note_title = "----Denied ") and InStr(full_note_title, "SNAP") <> 0 Then	'and if the CNOTE was of the approval
						Call write_value_and_transmit("V", note_row, 3)			'open the CNOTE detail and read the time it was created
						EMReadScreen note_time, 5, 9, 30
						' MsgBox note_time
						Do														'back out of the CNOTE DUMP information'
							PF3
							EMReadScreen still_in_dump, 4, 1, 48
						Loop until still_in_dump <> "DUMP"
					End If
				End If
				note_row = note_row + 1											'moving down the list of CNOTEs'
				if note_row = 19 then
					note_row = 5
					PF8
					EMReadScreen check_for_last_page, 9, 24, 14
					If check_for_last_page = "LAST PAGE" Then Exit Do
				End If
				EMReadScreen next_note_date, 8, note_row, 6
				if next_note_date = "        " then Exit Do
				' MsgBox next_note_date
				next_note_date = DateAdd("d", 0, next_note_date)
				' MsgBox "approved_date - " & approved_date & vbCr &  "too_old_date - " & too_old_date & vbCr & "next_note_date - " & next_note_date
			Loop until DateDiff("d", too_old_date, next_note_date) <= 0

			approval_date_and_time = approved_date & " " & note_time			'formatting the information found in ELIG and CNOTE'
			approval_date_and_time = trim(approval_date_and_time)
			approval_date_and_time = DateAdd("d", 0, approval_date_and_time)

			ObjReportExcel.Cells(total_excel_row, exch_app_date_time_col).Value = approval_date_and_time		'add the information to the worklist excel
			ObjReportExcel.Cells(total_excel_row, exch_app_status_col).Value = elig_status
			If expedited_status = "" Then ObjReportExcel.Cells(total_excel_row, exch_app_exp_status_col).Value = "False"
			If expedited_status = "EXPEDITED" Then ObjReportExcel.Cells(total_excel_row, exch_app_exp_status_col).Value = "True"
		End If
	End If

	total_excel_row = total_excel_row + 1
	MAXIS_case_number = trim(ObjReportExcel.Cells(total_excel_row, work_case_nbr_col_const).Value)
Loop until MAXIS_case_number = ""
Call back_to_SELF
objReportWorkbook.Save()		'saving the excel

If leave_excel_open = "No - Close the file" Then		'if the file should be closed - it does it here.
	ObjReportExcel.ActiveWorkbook.Close

	ObjReportExcel.Application.Quit
	ObjReportExcel.Quit
End If

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
const hss_rept_hss_name_col				= 7
const hss_rept_hss_email_col			= 8
const hss_rept_pm_name_col				= 9
const hss_rept_pm_email_col				= 10
const hss_rept_total_report_row_col		= 11

const worker_col 	= 1
const hss_name_col 	= 2
const hss_email_col = 3
const pm_name_col 	= 4
const pm_email_col 	= 5

const hsr_name_const 	= 0
const hss_name_const 	= 1
const hsr_email_const 	= 2
const pm_name_const 	= 3
const pm_email_const 	= 4

'This is where we get the information about HSRs and HSSs - we need to determine the data source and update this functionality once received - currently it is using a sheet in the HSS report out Excel File
Dim WORKER_ARRAY()
ReDim WORKER_ARRAY(pm_email_const, 0)

ObjHSSExcel.worksheets("Worker List").Activate

excel_row = 2
worker_count = 0
Do
	ReDim preserve WORKER_ARRAY(pm_email_const, worker_count)

	WORKER_ARRAY(hsr_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, worker_col))
	WORKER_ARRAY(hss_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, hss_name_col))
	WORKER_ARRAY(hsr_email_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, hss_email_col))
	WORKER_ARRAY(pm_name_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, pm_name_col))
	WORKER_ARRAY(pm_email_const, worker_count) = trim(ObjHSSExcel.Cells(excel_row, pm_email_col))

	excel_row = excel_row + 1
	worker_count = worker_count + 1
	next_worker_info = trim(ObjHSSExcel.Cells(excel_row, worker_col))
Loop until next_worker_info = ""

'Now we are going to read all of the txt files generated from the Expedited Determination script runs. If they need review, they will be added to the report out file - otherwise they will be saved in the large Excel list
this_month_worklist = MonthName(DatePart("m", date)) & " " & DatePart("yyyy", date)			'finding the right sheet for the HSS report out excel'
ObjHSSExcel.worksheets(this_month_worklist).Activate							'activating the sheet

'finding the last row of the HSS Report Out ffile
hss_excel_row = 1																'default to the first row
Do
	hss_excel_row = hss_excel_row + 1
	this_case_number = trim(ObjHSSExcel.Cells(hss_excel_row, 2).Value)
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
					If line_info(0) = "WORKER NAME" Then
						ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_col).Value = line_info(1)
						For each_wrkr = 0 to UBound(WORKER_ARRAY, 2)			'here we need to use the data of HSRs and HSSs to fill in the appropriate HSS and PM based on Worker Name
							If WORKER_ARRAY(hsr_name_const, each_wrkr) = line_info(1) Then
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_name_col).Value = WORKER_ARRAY(hss_name_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_email_col).Value = WORKER_ARRAY(hsr_email_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_name_col).Value = WORKER_ARRAY(pm_name_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_email_col).Value = WORKER_ARRAY(pm_email_const, each_wrkr)
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_total_report_row_col).Value = total_excel_row
								ObjHSSExcel.Cells(hss_excel_row, hss_rept_report_day_col).Value = date
							End If
						Next
					End If

				End If
			Next
			hss_excel_row = hss_excel_row + 1
		End If

		total_excel_row = total_excel_row + 1										'advance to the next row
		STATS_counter = STATS_counter + 1

		objTextStream.Close						'we are done with this file, so we must close the access
		objFSO.MoveFile this_file_path , txt_file_archive_path & "\" & this_file_name & ".txt"    'moving each file to the archive file
	End If
Next

objWorkbook.Save()		'saving the excel
objHSSWorkbook.Save()		'saving the excel

''THIS PART is the Emailing
'Commented out as we wait for data and emailing authority

'constants for the report out array of cases that need to be reported to HSSs'
const case_numb_rept_out_const		= 0
const app_date_rept_out_const		= 1
const intv_date_rept_out_const		= 2
const delay_explain_rept_out_const	= 3
const worker_name_rept_out_const	= 4
const hss_name_rept_out_const		= 5
const hss_email_rept_out_const		= 6
const pm_name_rept_out_const		= 7
const pm_email_rept_out_const		= 8
const last_rept_out_const			= 9

'defining the array'
Dim REPORT_OUT_ARRAY()
ReDim REPORT_OUT_ARRAY(last_rept_out_const, 0)

all_hsr_list = "~|~"															'here we set lists to make sure we do not duplicate any people
all_hss_list = "~"
all_pm_list = "~"

hss_excel_row = 2																'default to the first row
report_out_count = 0
Do
	If ObjHSSExcel.Cells(hss_excel_row, hss_rept_report_day_col).Value = date Then	'only ffinding cases from today's report
		ReDim preserve REPORT_OUT_ARRAY(last_rept_out_const, report_out_count)	''resize the array

		REPORT_OUT_ARRAY(case_numb_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_case_number_col).Value					'setting the excel infomraiton into the array
		REPORT_OUT_ARRAY(app_date_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_application_date_col).Value
		REPORT_OUT_ARRAY(intv_date_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_interview_date_col).Value
		REPORT_OUT_ARRAY(delay_explain_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_approval_delay_detail_col).Value
		REPORT_OUT_ARRAY(worker_name_rept_out_const, report_out_count) = trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_script_user_col).Value)
		REPORT_OUT_ARRAY(hss_name_rept_out_const, report_out_count) = trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_name_col).Value)
		REPORT_OUT_ARRAY(hss_email_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_hss_email_col).Value & "@hennepin.us"
		REPORT_OUT_ARRAY(pm_name_rept_out_const, report_out_count) = trim(ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_name_col).Value)
		REPORT_OUT_ARRAY(pm_email_rept_out_const, report_out_count) = ObjHSSExcel.Cells(hss_excel_row, hss_rept_pm_email_col).Value & "@hennepin.us"

		'determining if the people are already in the list or not
		If InStr(all_hsr_list, "~" & REPORT_OUT_ARRAY(worker_name_rept_out_const, report_out_count) & "|" & REPORT_OUT_ARRAY(hss_name_rept_out_const, report_out_count) & "~")  = 0 Then all_hsr_list = all_hsr_list & REPORT_OUT_ARRAY(worker_name_rept_out_const, report_out_count) & "|" & REPORT_OUT_ARRAY(hss_name_rept_out_const, report_out_count) & "~"
		If InStr(all_hss_list, "~" & REPORT_OUT_ARRAY(hss_name_rept_out_const, report_out_count) & "~") = 0 Then all_hss_list = all_hss_list & REPORT_OUT_ARRAY(hss_name_rept_out_const, report_out_count) & "~"
		If InStr(all_pm_list, "~" & REPORT_OUT_ARRAY(pm_name_rept_out_const, report_out_count) & "~") = 0 Then all_pm_list = all_pm_list & REPORT_OUT_ARRAY(pm_name_rept_out_const, report_out_count) & "~"

		report_out_count = report_out_count + 1
	end If
	hss_excel_row = hss_excel_row + 1
	this_case_number = trim(ObjHSSExcel.Cells(hss_excel_row, 2).Value)
Loop Until this_case_number = ""

HSR_ARRAY = split(all_hsr_list, "~")		'creating an array of all the people from the report
HSS_ARRAY = split(all_hss_list, "~")
' PM_ARRAY = split(all_pm_list, "~")

all_repts_list = ""							'this is a string that will be used to add all the detail of the report out to the emails
email_subject = "Cases determined EXPEDITED that Require Action"

For each_hss = 1 to UBound(HSS_ARRAY)-1		'there is always a blank instance at the begining and end so we start and the second and end at the second to last
	report_details = ""						'this string will add details of the report out to each email by HSS

	email_recip = ""						'blanking out variable for each loop through the HSSs
	email_recip_CC = ""
	email_name = ""
	For each_hsr = 1 to UBound(HSR_ARRAY)-1			'now we check each HSR on our list
		temp_array = ""
		temp_array = split(HSR_ARRAY(each_hsr), "|")			'The HSR information has the HSR name and the HSS name associated with that HSR divided by a line- defined when creating the HSR string on line 573
		If temp_array(1) =  HSS_ARRAY(each_hss) Then			'If the HSS for this HSR matches the HSS we are corrrently emailing, then we are foing to add this information to the email
			report_details = report_details & vbCr & vbCr & "Cases found processed by: " & temp_array(0)		'Adding the HSR name to the email
			For each_rept = 0 to UBound(REPORT_OUT_ARRAY, 2)													'Now we look at each report item and if the HSR matches - the case detail is added to the email string
				If REPORT_OUT_ARRAY(hss_name_rept_out_const, each_rept) = HSS_ARRAY(each_hss) and REPORT_OUT_ARRAY(worker_name_rept_out_const, each_rept) = temp_array(0) Then
					report_details = report_details & vbCr & " - " & REPORT_OUT_ARRAY(case_numb_rept_out_const, each_rept) & " Application Date: " & REPORT_OUT_ARRAY(app_date_rept_out_const, each_rept) & " Interview Date: " & REPORT_OUT_ARRAY(intv_date_rept_out_const, each_rept)
					If REPORT_OUT_ARRAY(delay_explain_rept_out_const, each_rept) <> "" Then report_details = report_details & vbCr & chr(9) & " - Explanation of Delay: " & REPORT_OUT_ARRAY(delay_explain_rept_out_const, each_rept)

					If email_name = "" Then email_name = REPORT_OUT_ARRAY(hss_name_rept_out_const, each_rept)			''setting the names and emails
					If email_recip = "" Then email_recip = REPORT_OUT_ARRAY(hss_email_rept_out_const, each_rept)
					If email_recip_CC = "" Then email_recip_CC = REPORT_OUT_ARRAY(pm_email_rept_out_const, each_rept)
				End if
			next
		End If
	Next

	'Now we write the email for this particular HSS
	email_body = "Good morning " & email_name & ", "
	email_body = email_body & vbCr & vbCr & "This is an automated email to provide a list of cases that require review and action. "
	email_body = email_body & vbCr & "The case(s) listed in this email were determined as eligibility for Expedited SNAP but are still in a PENDING status. The case(s) were worked on yesterday and action was likely required at that time. Reach out to the worker(s) and ensure they have the necessary support to complete processing on these case(s) today."
	email_body = email_body & report_details
	email_body = email_body & vbCr & vbCr & "Cases that meet expedited criteria can have all verifications except identity of the applicant postponed. No other verifications should hold up processing of Expedited SNAP (this included immigration verification - do not hold cases us for immigration verification)."
	email_body = email_body & vbCr & "The only other instances in which we cannot approve expedited right away is in the case of a resident still in a facility or if the last issueance of SNAP was Expedited with postponed verifications and there are currently postponed verifications. "
	email_body = email_body & vbCr & "If the worker beleives the case cannot be processed at this time, ensure they check with Knowledge Now. Any other policy or procedural questions should also go to Knowledge Now."
	email_body = email_body & vbCr & vbCr & "*** Remember cases that are Expedited do NOT have a 30 Day application processing period, they have a 5 Business Day/7 Calendar Day application processing period. ***"
	email_body = email_body & vbCr & vbCr & "Please connect with QI Leadership or the BlueZone Script Team with any questions about this report."
	email_body = email_body & vbCr & vbCr & "Thank you for your dedication to our residents and quality processing."

	Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, "", False)
	'adding this email detail to the large email with all the reportout information
	all_repts_list = all_repts_list & vbCr  & "______________________________________________________________________________" & vbCr & "Sent to: " & email_name & " - " & email_recip & vbCr & "CC: " & email_recip_CC & report_details
Next
'here we send the large all report email
email_body = "These are all the cases identified by the Expedited Determination Report Process"
email_body = email_body & vbCr & vbCr & "This is an automated email to provide a list of cases that require review and action. "
email_body = email_body & vbCr & "The case(s) listed in this email were determined as eligibility for Expedited SNAP but are still in a PENDING status. The case(s) were worked on yesterday and action was likely required at that time. Reach out to the worker(s) and ensure they have the necessary support to complete processing on these case(s) today."
email_body = email_body & vbCr & all_repts_list
Call create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", email_subject, email_body, "", False)

If leave_excel_open = "No - Close the file" Then
	ObjHSSExcel.ActiveWorkbook.Close

	ObjHSSExcel.Application.Quit
	ObjHSSExcel.Quit

	objExcel.ActiveWorkbook.Close

	objExcel.Application.Quit
	objExcel.Quit
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
