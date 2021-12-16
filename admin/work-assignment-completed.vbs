'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Work Assignment Completed.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 720          'manual run time in seconds
STATS_denomination = "I"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("07/08/2020", "Added message reminding you to let the script work without trying to multitask. This can cause the script to error.", "Casey Love, Hennepin County")
call changelog_update("06/01/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
Function File_Exists(file_name, does_file_exist)
    ' Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(file_name)) Then
        does_file_exist = True
    Else
      does_file_exist = False
    End If
End Function
'DECLARATIONS ==============================================================================================================

Const date_col                          = 1                         'These are the constants for the columns in the tracking Excel documents
Const worker_id_col                     = 2
Const worker_name_col                   = 3
Const cases_reviewed_col                = 4
Const cases_xfs_app_col                 = 5
Const cases_xfs_no_id_col               = 6
Const cases_xfs_id_app                  = 7
Const cases_xfs_correct_col             = 8
Const cases_xfs_no_caf                  = 9
Const cases_xfs_verifs_not_postponed_col = 10
Const cases_xfs_MAXIS_wrong_col         = 11
Const cases_xfs_bad_note_col            = 12
Const xfs_assignment_length_col         = 13
Const xfs_assignment_assessment_col     = 14
Const xfs_list_of_cases_col             = 15
Const other_notes_col					= 16

Const cases_d30_no_interview            = 5
Const cases_d30_other_reason            = 6
Const cases_d30_app                     = 7
Const cases_d30_timely                  = 8
Const cases_d30_not_timely              = 9
Const cases_d30_future_verifs           = 10
Const d30_assignment_length_col         = 11
Const d30_assignment_assessment_col     = 12
Const d30_list_of_cases_col             = 13

Dim TABLE_ARRAY()

'SCRIPT ====================================================================================================================
'Find who is running
Set objNet = CreateObject("WScript.NetWork")                                    'getting the users windows ID
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

For each tester in tester_array                                                 'Loop through all the testers in the array to see if the user is in the list of testers.
    If user_ID_for_validation = tester.tester_id_number Then
        qi_worker_full_name            = tester.tester_full_name
        qi_worker_first_name           = tester.tester_first_name
        qi_worker_last_name            = tester.tester_last_name
        qi_worker_email                = tester.tester_email
        qi_worker_id_number            = tester.tester_id_number
        qi_worker_x_number             = tester.tester_x_number
        qi_worker_supervisor           = tester.tester_supervisor_name
        qi_worker_supervisor_email     = tester.tester_supervisor_email
        qi_worker_test_groups          = tester.tester_groups
        qi_staff = FALSE
        For each group in qi_worker_test_groups                                 'looking at all of the groups this tester is a part of to see if QI or BZ
            If group = "QI" Then qi_staff = TRUE
            If group = "BZ" Then qi_staff = TRUE
        Next
    End If
Next
'If this did not find the user is a tester for QI the script will end as this is only for QI staff - access to the files and folders will be restricted and the script will fail
If qi_staff = FALSE Then script_end_procedure_with_error_report("This script is for QI specific processes and only for QI staff. You are not listed as QI staff and running this script could cause errors in data reccording and QI processes. Please contact the BlueZone script team or pres 'Yes' below if you believe this to be in error.")

work_assignment_date = date & ""                'defaulting some of the variables for the initial values
email_signature = qi_worker_first_name
type_of_work_assignment = "On Demand Applications"

'Dialog to determine who you are and what kind of assignment you finished.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 306, 110, "Work Assingment Selection"
  EditBox 150, 5, 150, 15, qi_worker_full_name
  EditBox 150, 25, 150, 15, work_assignment_date
  ' DropListBox 150, 45, 150, 45, "Select One..."+chr(9)+"Expedited Review"+chr(9)+"Pending at Day 30 - Part of On Demand", type_of_work_assignment
  DropListBox 150, 45, 150, 45, "Select One..."+chr(9)+"On Demand Applications", type_of_work_assignment
  EditBox 150, 65, 150, 15, email_signature
  ButtonGroup ButtonPressed
    OkButton 195, 90, 50, 15
    CancelButton 250, 90, 50, 15
	PushButton 5, 95, 70, 10, "INSTRUCTIONS", instructions_btn
  Text 20, 10, 125, 10, " QI Staff Member completing the work:"
  Text 10, 30, 140, 10, "Date of assignment and work completion:"
  Text 85, 50, 60, 10, "Assignment Type:"
  Text 80, 70, 65, 10, "   Sign your emails:"
EndDialog

Do
    err_msg = ""

    dialog Dialog1
    cancel_without_confirmation

    qi_worker_full_name = trim(qi_worker_full_name)
    email_signature = trim(email_signature)

    'Everything is required in this dialog.
    If qi_worker_full_name = "" Then err_msg = err_msg & vbNewLine & "* Enter your full name."
    If IsDate(work_assignment_date) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter the date of the assignment, the day you worked on the assignment list."
    If type_of_work_assignment = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select which work assignment you are completing for the day."
    If email_signature = "" Then err_msg = err_msg & vbNewLine & "* Enter how you want your email signed."

	If ButtonPressed = instructions_btn Then
		Call open_URL_in_browser("https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/ADMIN/ADMIN%20-%20WORK%20ASSIGNMENT%20COMPLETED.docx")
		err_msg = "LOOP" & err_msg
    ElseIf err_msg <> "" Then
		MsgBox "Please resolve to continue:" & vbNewLine & err_msg
	End If

Loop until err_msg = ""



work_assignment_date = FormatDateTime(work_assignment_date, 2)                  'date formats to be sure the year is 4 digits
month_of_assignment = right("0" & DatePart("m", work_assignment_date), 2)       'Pulling the month and year of the assignment for use in doc names and folders.
year_of_assignment = DatePart("yyyy", work_assignment_date)
date_for_doc = work_assignment_date & ""
date_for_doc = replace(date_for_doc, "/", "-")                                  'taking the '/' out for the doc names because otherwise it can't save

word_doc_name = ""
word_doc_file_path = ""

'Dialog to gather the details/stats/counts
Select Case type_of_work_assignment                                             'differen selections/options based on the work assignment selection
	Case "On Demand Applications"
		close_worklist_msgbox = MsgBox("This script can only function properly if On Demand Daily Worklist is saved and closed. Be sure you have finished your notes and entered all informationn on the worklist, save it and closed the file." & vbCr & vbCr & "Do it now if you have it open." & vbCr & vbCr & "Is the worklist for this assignment closed?", vbQuestion + vbYesNo, "Close the worklist")
		If close_worklist_msgbox = vbNo Then script_end_procedure("Complete the work on the list, then save and close the file. Once that is done you can rerun this script to capture the end of the work assignment. This script will now end.")

		file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & date_for_doc & " Worklist.xlsx"

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 556, 235, "Details of Cases Pending over 30 Days Work Assignment"
		  EditBox 80, 40, 30, 15, oda_number_of_cases_reviewed
		  EditBox 125, 60, 30, 15, oda_number_of_cases_denied_no_interview
		  EditBox 130, 80, 30, 15, oda_number_of_appt_notc
		  EditBox 115, 100, 30, 15, oda_number_of_nomis
		  EditBox 135, 120, 30, 15, oda_number_correction_emails
		  EditBox 165, 140, 30, 15, oda_number_prog_updated
		  EditBox 125, 160, 30, 15, oda_number_of_case_notes
		  EditBox 405, 35, 20, 15, assignment_hours
		  EditBox 455, 35, 20, 15, assignment_minutes
		  DropListBox 400, 60, 110, 45, "Select One..."+chr(9)+"Great"+chr(9)+"Good"+chr(9)+"Okay"+chr(9)+"Neutral"+chr(9)+"A little rough"+chr(9)+"Bad"+chr(9)+"Terrible", assignment_assesment
		  EditBox 265, 95, 285, 15, assignment_case_numbers_to_save
		  EditBox 265, 135, 285, 15, assignment_new_ideas
		  EditBox 10, 195, 540, 15, assignment_other_notes
		  ButtonGroup ButtonPressed
		    OkButton 450, 215, 50, 15
		    CancelButton 500, 215, 50, 15
		  Text 115, 10, 295, 10, "****************** On Demand Waiver Applications - QI Daily Assignment ******************"
		  GroupBox 10, 25, 245, 155, "Number of cases that:"
		  Text 20, 45, 60, 10, "... you reviewed:"
		  Text 20, 65, 100, 10, ".. you denied for no interview:"
		  Text 20, 85, 110, 10, ".. you sent a manual APPT NOTC:"
		  Text 20, 105, 90, 10, ".. you sent a manual NOMI:"
		  Text 20, 125, 110, 10, ".. you sent a correction email on:"
		  Text 20, 145, 145, 10, ".. you updated PROG with an interview date:"
		  Text 20, 165, 105, 10, ".. you added a CASE:NOTE on:"
		  Text 265, 40, 140, 10, "About how long did the assignment take?"
		  Text 430, 40, 20, 10, "hours"
		  Text 480, 40, 30, 10, "minutes"
		  Text 265, 65, 135, 10, "How was the assignment for you today?"
		  Text 265, 85, 180, 10, "Any case numbers to save for example/larger reivew?"
		  Text 265, 125, 105, 10, "Ideas of other counts to collect:"
		  Text 265, 150, 285, 15, "These are common errors or handling that we are seeing in review, this would be to add to the option on the left."
		  Text 10, 185, 140, 10, "Other notes about assignment from today:"
		EndDialog

		Do
			Do
				err_msg = ""

				dialog Dialog1
				cancel_confirmation

				EditBox 80, 40, 30, 15, oda_number_of_cases_reviewed
	  		  EditBox 125, 60, 30, 15, oda_number_of_cases_denied_no_interview
	  		  EditBox 130, 80, 30, 15, oda_number_of_appt_notc
	  		  EditBox 115, 100, 30, 15, oda_number_of_nomis
	  		  EditBox 135, 120, 30, 15, oda_number_correction_emails
	  		  EditBox 165, 140, 30, 15, oda_number_prog_updated
	  		  EditBox 125, 160, 30, 15, oda_number_of_case_notes
				If IsNumeric(oda_number_of_cases_reviewed) = FALSE OR IsNumeric(oda_number_of_cases_denied_no_interview) = FALSE OR IsNumeric(oda_number_of_appt_notc) = FALSE OR IsNumeric(oda_number_of_nomis) = FALSE OR IsNumeric(oda_number_correction_emails) = FALSE OR IsNumeric(oda_number_prog_updated) = FALSE OR IsNumeric(oda_number_of_case_notes) = FALSE Then
					err_msg = err_msg & vbNewLine & "* Count needed. Enter the number of cases that meet the following criteria: "
					If IsNumeric(oda_number_of_cases_reviewed) = FALSE Then err_msg = err_msg & vbNewLine & "  - total you reviewed (1st)"
					If IsNumeric(oda_number_of_cases_denied_no_interview) = FALSE Then err_msg = err_msg & vbNewLine & "  - you denied for no interview (2nd)"
					If IsNumeric(oda_number_of_appt_notc) = FALSE Then err_msg = err_msg & vbNewLine & "  - you sent a manual appoinntment notice (3rd)"
					If IsNumeric(oda_number_of_nomis) = FALSE Then err_msg = err_msg & vbNewLine & "  - you sent a manual nomi (4th)"
					If IsNumeric(oda_number_correction_emails) = FALSE Then err_msg = err_msg & vbNewLine & "  - you sent a correction email (5th)"
					If IsNumeric(oda_number_prog_updated) = FALSE Then err_msg = err_msg & vbNewLine & "  - you updated PROG with the interview date (6th)"
					If IsNumeric(oda_number_of_case_notes) = FALSE Then err_msg = err_msg & vbNewLine & "  - you entered a CASE:NOTE (7th)"
					err_msg = err_msg & vbNewLine & "  These counts are crucial to tracking our progress and advancement in our work assignment efforts." & vbNewLine
				End If

				If IsNumeric(assignment_hours) = FALSE AND IsNumeric(assignment_minutes) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter how long it took for you to complete the assignment. This can be entered in hours and/or minutes. Do your best to discount breaks and other work so we get a good idea of how long the work is taking."
				If assignment_assesment = "Select One..." Then err_msg = err_msg & vbNewLine & "* Let us know how the work was for you today."

				If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = FALSE

		work_assignment_date = FormatDateTime(work_assignment_date, 2)                  'date formats to be sure the year is 4 digits
		month_of_assignment = right("0" & DatePart("m", work_assignment_date), 2)       'Pulling the month and year of the assignment for use in doc names and folders.
		year_of_assignment = DatePart("yyyy", work_assignment_date)
		date_for_doc = work_assignment_date & ""
		date_for_doc = replace(date_for_doc, "/", "-")                                  'taking the '/' out for the doc names because otherwise it can't save

		today_file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/QI " & date_for_doc & " Worklist.xlsx"
		archive_file_selection_path = t_drive & "/Eligibility Support/Restricted/QI - Quality Improvement/REPORTS/On Demand Waiver/QI On Demand Daily Assignment/Archive/" & year_of_assignment & "-" & month_of_assignment & "QI " & date_for_doc & " Worklist.xlsx"

		Call File_Exists(today_file_selection_path, does_file_exist)
		If does_file_exist = True Then file_selection_path = today_file_selection_path
		If does_file_exist = False Then file_selection_path = archive_file_selection_path

		If IsNumeric(assignment_minutes) = FALSE Then assignment_minutes = 0    'creating a time variable with ONLY minutes for the spreadsheet
		If IsNumeric(assignment_hours) = TRUE Then
			minutes_from_hours = assignment_hours * 60
			assignment_time = assignment_minutes + minutes_from_hours
		Else
			assignment_time = assignment_minutes
		End If
		case_list_sheet_name = "Work List for " & date_for_doc

		call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)

		ObjExcel.Worksheets("Statistics").visible = True
		ObjExcel.worksheets("Statistics").Activate
		ObjExcel.Cells(2, 5).Value = qi_worker_full_name
		ObjExcel.Cells(3, 5).Value = date
		ObjExcel.Cells(4, 5).Value = time
		ObjExcel.Cells(5, 5).Value = oda_number_of_cases_reviewed
		ObjExcel.Cells(6, 5).Value = oda_number_of_cases_denied_no_interview
		ObjExcel.Cells(7, 5).Value = oda_number_prog_updated
		ObjExcel.Cells(8, 5).Value = oda_number_of_appt_notc
		ObjExcel.Cells(9, 5).Value = oda_number_of_nomis
		ObjExcel.Cells(10, 5).Value = oda_number_correction_emails
		ObjExcel.Cells(11, 5).Value = oda_number_of_case_notes
		ObjExcel.Cells(12, 5).Value = assignment_time
		ObjExcel.Cells(13, 5).Value = assignment_assesment
		ObjExcel.Cells(14, 5).Value = assignment_case_numbers_to_save
		ObjExcel.Cells(15, 5).Value = assignment_new_ideas
		ObjExcel.Cells(16, 5).Value = assignment_other_notes

		ObjExcel.worksheets(case_list_sheet_name).Activate
		ObjExcel.Worksheets("Statistics").visible = False

		objWorkbook.Save
		ObjExcel.Quit

		excel_file_path = file_selection_path

		main_email_body = "The On Demand Appplication Assignment has been completed for " & work_assignment_date & "."
		main_email_subject = "On Demand Application worklist Complete"

	Case "Expedited Review"
		file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\QI Expedited Review " & date_for_doc & ".xlsx" 'single assignment file

		Set fso = CreateObject("Scripting.FileSystemObject")

		If (fso.FileExists(file_selection_path)) Then
			Call excel_open(file_selection_path, True, True, ObjWorkExcel, objWorkbook)  'opens the selected excel file'

			exp_number_of_cases_reviewed = 0
			exp_number_of_cases_approved = 0
			exp_number_of_cases_no_id = 0
			exp_number_of_cases_ID_approved = 0
			exp_number_of_cases_correct = 0
			exp_number_of_cases_xfs_no_caf = 0
			exp_number_of_cases_should_have_been_postponed = 0
			exp_number_of_cases_MAXIS_incorrect = 0
			exp_number_of_cases_bad_notes = 0

			excel_row = 2
			Do
				the_case_number = trim(ObjWorkExcel.Cells(excel_row, 2).Value)

				If the_case_number <> "" Then
					If ObjWorkExcel.Cells(excel_row,  9).Value = "Yes" Then exp_number_of_cases_reviewed = exp_number_of_cases_reviewed + 1 									'Rewiewed
					If ObjWorkExcel.Cells(excel_row, 10).Value = "Yes" Then exp_number_of_cases_approved = exp_number_of_cases_approved + 1 									'Approved
					If ObjWorkExcel.Cells(excel_row, 11).Value = "Yes" Then exp_number_of_cases_no_id = exp_number_of_cases_no_id + 1 											'Appear EXP, no ID, could not be approved
					If ObjWorkExcel.Cells(excel_row, 12).Value = "Yes" Then exp_number_of_cases_ID_approved = exp_number_of_cases_ID_approved + 1 								'Appear EXP, ID was available - Incorrect
					If ObjWorkExcel.Cells(excel_row, 13).Value = "Yes" Then exp_number_of_cases_correct = exp_number_of_cases_correct + 1 										'Processed correctly by HSR
					If ObjWorkExcel.Cells(excel_row, 14).Value = "Yes" Then exp_number_of_cases_xfs_no_caf = exp_number_of_cases_xfs_no_caf + 1 								'No CAF on file
					If ObjWorkExcel.Cells(excel_row, 15).Value = "Yes" Then exp_number_of_cases_should_have_been_postponed = exp_number_of_cases_should_have_been_postponed + 1 'Verifications should have been postponed/Case app'd
					If ObjWorkExcel.Cells(excel_row, 16).Value = "Yes" Then exp_number_of_cases_MAXIS_incorrect = exp_number_of_cases_MAXIS_incorrect + 1 						'MAXIS was updated incorrectly
					If ObjWorkExcel.Cells(excel_row, 17).Value = "Yes" Then exp_number_of_cases_bad_notes = exp_number_of_cases_bad_notes + 1 									'Has insufficient CASE/NOTES
					If ObjWorkExcel.Cells(excel_row, 18).Value = "Yes" Then assignment_case_numbers_to_save = assignment_case_numbers_to_save & ", " & the_case_number 			'Save case number for team review?
				End If

				excel_row = excel_row + 1
			Loop Until the_case_number = ""

			ObjWorkExcel.ActiveWorkbook.Close
			ObjWorkExcel.Application.Quit
			ObjWorkExcel.Quit
		End If

		If left(assignment_case_numbers_to_save, 2) = ", " Then assignment_case_numbers_to_save = right(assignment_case_numbers_to_save, len(assignment_case_numbers_to_save) - 2)
		exp_number_of_cases_reviewed = exp_number_of_cases_reviewed & ""
		exp_number_of_cases_approved = exp_number_of_cases_approved & ""
		exp_number_of_cases_no_id = exp_number_of_cases_no_id & ""
		exp_number_of_cases_ID_approved = exp_number_of_cases_ID_approved & ""
		exp_number_of_cases_correct = exp_number_of_cases_correct & ""
		exp_number_of_cases_xfs_no_caf = exp_number_of_cases_xfs_no_caf & ""
		exp_number_of_cases_should_have_been_postponed = exp_number_of_cases_should_have_been_postponed & ""
		exp_number_of_cases_MAXIS_incorrect = exp_number_of_cases_MAXIS_incorrect & ""
		exp_number_of_cases_bad_notes = exp_number_of_cases_bad_notes & ""

        counts_number = 8                                                       'There are 9 counts we select so the array goes to 8
        'Setting the file locations and doc strings
        word_doc_name = qi_worker_full_name & " - EXP Processing Assignment Report for " & date_for_doc
        word_doc_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & month_of_assignment &"-" & year_of_assignment & "\Report Out\"
        excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP Work Counts.xlsx"

        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 586, 330, "Details of Expedited Processing Work Assignment"
          EditBox 80, 40, 30, 15, exp_number_of_cases_reviewed
          EditBox 80, 60, 30, 15, exp_number_of_cases_approved
          EditBox 220, 80, 30, 15, exp_number_of_cases_no_id
          EditBox 175, 110, 30, 15, exp_number_of_cases_ID_approved
          EditBox 155, 145, 30, 15, exp_number_of_cases_correct
          EditBox 100, 165, 30, 15, exp_number_of_cases_xfs_no_caf
          EditBox 205, 185, 30, 15, exp_number_of_cases_should_have_been_postponed
          EditBox 135, 215, 30, 15, exp_number_of_cases_MAXIS_incorrect
          EditBox 205, 235, 30, 15, exp_number_of_cases_bad_notes
          Text 145, 10, 315, 10, "********* Expedited Processing - the review of the daily assignment of Expedited SNAP *********"
          GroupBox 10, 25, 245, 250, "Number of cases that:"
          Text 20, 45, 60, 10, ".. you reviewed:"
          Text 20, 65, 55, 10, ".. you approved:"
          Text 20, 85, 195, 10, ".. appear expedited but have no ID, could not be approved:"
          Text 30, 95, 160, 10, "These are cases that are correct in waiting on ID"
          Text 20, 115, 150, 10, ".. appear expedited and an ID WAS available:"
          Text 30, 125, 140, 20, "These are cases that were delayed by the HSR but should have been approved."
          Text 20, 150, 130, 10, ".. were processed correctly by the HSR:"
          Text 20, 170, 75, 10, ".. have no CAF on file:"
          Text 20, 190, 180, 10, ".. have verifications that should have been postponed:"
          Text 35, 200, 170, 10, "Cases that could have been approved but were not."
          Text 20, 220, 115, 10, ".. MAXIS was updated incorrectly:"
          Text 20, 240, 180, 10, ".. have insufficient CASE:NOTEs about the application:"
          Text 30, 250, 170, 20, "These include cases where scripts were not used and the information was not provided manually."
          EditBox 405, 35, 20, 15, assignment_hours
          EditBox 455, 35, 20, 15, assignment_minutes
          DropListBox 400, 60, 110, 45, "Select One..."+chr(9)+"Great"+chr(9)+"Good"+chr(9)+"Okay"+chr(9)+"Neutral"+chr(9)+"A little rough"+chr(9)+"Bad"+chr(9)+"Terrible", assignment_assesment
          EditBox 265, 95, 285, 15, assignment_case_numbers_to_save
          EditBox 265, 135, 285, 15, assignment_new_ideas
          EditBox 10, 290, 570, 15, assignment_other_notes
          ButtonGroup ButtonPressed
            OkButton 480, 310, 50, 15
            CancelButton 530, 310, 50, 15
          Text 265, 40, 140, 10, "About how long did the assignment take?"
          Text 430, 40, 20, 10, "hours"
          Text 480, 40, 30, 10, "minutes"
          Text 265, 65, 135, 10, "How was the assignment for you today?"
          Text 265, 85, 180, 10, "Any case numbers to save for example/larger reivew?"
          Text 265, 125, 105, 10, "Ideas of other counts to collect:"
          Text 265, 150, 285, 15, "These are common erros or handling that we are seeing in review, this would be to add to the option on the left."
          Text 10, 280, 140, 10, "Other notes about assignment from today:"
          Text 265, 215, 305, 10, "COUNTS should be based on the discovery YOU have made today."
          Text 265, 230, 290, 25, "We can only get accurate data if we are not duplicating the case counts. If the notes in the assignment spreadsheet that one of the count criteria was met, we must trust that worker counted it when they did the initial discovery."
          Text 265, 265, 260, 15, "The only exception is the number of cases reviewed, enter the total number of cases assigned to review that you checked."
        EndDialog

        Do
            Do
                err_msg = ""

                dialog Dialog1
                cancel_confirmation

                'All the counts are required.
                If IsNumeric(exp_number_of_cases_reviewed) = FALSE OR IsNumeric(exp_number_of_cases_approved) = FALSE OR IsNumeric(exp_number_of_cases_no_id) = FALSE OR IsNumeric(exp_number_of_cases_ID_approved) = FALSE OR IsNumeric(exp_number_of_cases_correct) = FALSE OR IsNumeric(exp_number_of_cases_xfs_no_caf) = FALSE OR IsNumeric(exp_number_of_cases_should_have_been_postponed) = FALSE OR IsNumeric(exp_number_of_cases_MAXIS_incorrect) = FALSE OR IsNumeric(exp_number_of_cases_bad_notes) = FALSE Then
                    err_msg = err_msg & vbNewLine & "* Count needed. Enter the number of cases that meet the following criteria: "
                    If IsNumeric(exp_number_of_cases_reviewed) = FALSE Then err_msg = err_msg & vbNewLine & "  - total you reviewed (1st)"
                    If IsNumeric(exp_number_of_cases_approved) = FALSE Then err_msg = err_msg & vbNewLine & "  - you approved (2nd)"
                    If IsNumeric(exp_number_of_cases_no_id) = FALSE Then err_msg = err_msg & vbNewLine & "  - with no identity document (3rd)"
                    If IsNumeric(exp_number_of_cases_ID_approved) = FALSE Then err_msg = err_msg & vbNewLine & "  - that actually has an ID and could be approved (4th)"
                    If IsNumeric(exp_number_of_cases_correct) = FALSE Then err_msg = err_msg & vbNewLine & "  - were processed correctly by HSR (5th)"
                    If IsNumeric(exp_number_of_cases_xfs_no_caf) = FALSE Then err_msg = err_msg & vbNewLine & "  - without a CAF (6th)"
                    If IsNumeric(exp_number_of_cases_should_have_been_postponed) = FALSE Then err_msg = err_msg & vbNewLine & "  - that have verifs that should have been postponed (7th)"
                    If IsNumeric(exp_number_of_cases_MAXIS_incorrect) = FALSE Then err_msg = err_msg & vbNewLine & "  - MAXIS was updated incorrectly (8th)"
                    If IsNumeric(exp_number_of_cases_bad_notes) = FALSE Then err_msg = err_msg & vbNewLine & "  - with insufficient CASE:NOTEs (9th)"
                    err_msg = err_msg & vbNewLine & "  These counts are crucial to tracking our progress and advancement in our work assignment efforts." & vbNewLine
                End If

                If IsNumeric(assignment_hours) = FALSE AND IsNumeric(assignment_minutes) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter how long it took for you to complete the assignment. This can be entered in hours and/or minutes. Do your best to discount breaks and other work so we get a good idea of how long the work is taking."
                If assignment_assesment = "Select One..." Then err_msg = err_msg & vbNewLine & "* Let us know how the work was for you today."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE

		leave_alone_msg = MsgBox("Thank you for providing all the detail." & vbNewLine & vbNewLine & "The script now needs to save all of the information, this can take a minute or two. The script needs to open Excel, Word, and send emails. This will work best if the script does not get interrupted." & vbNewLine & vbNewLine & "Do not touch the computer until the script finishes.", vbExclamation + vbSystemModal, "Let the Script Work Uninterrupted")

        If IsNumeric(assignment_minutes) = FALSE Then assignment_minutes = 0    'creating a time variable with ONLY minutes for the spreadsheet
        If IsNumeric(assignment_hours) = TRUE Then
            minutes_from_hours = assignment_hours * 60
            assignment_time = assignment_minutes + minutes_from_hours
        Else
            assignment_time = assignment_minutes
        End If

        call excel_open(excel_file_path, False,  False, ObjExcel, objWorkbook)  'opening the EXP SNAP Assignment worksheet

        sheet_selection = month_of_assignment &"-" & year_of_assignment         'The sheets are nameed MM-YYYY - this will use the date of the assignment to select the right sheet and open it.
        ObjExcel.worksheets(sheet_selection).Activate

        excel_row = 2                                                           'Finding the first open Excel Row
        Do
            this_entry = ObjExcel.Cells(excel_row, 1).Value
            this_entry = trim(this_entry)
            If this_entry <> "" Then excel_row = excel_row + 1
        Loop until this_entry = ""
		assignment_other_notes = trim(assignment_other_notes)
		assignment_other_notes = " - Completed on: " & date & " " & time & " - " & assignment_other_notes
		If right(assignment_other_notes, 3) = " - " Then assignment_other_notes = left(assignment_other_notes, len(assignment_other_notes) - 3)

        'Adding the information from the dialog into the Excel spreadsheet
        ObjExcel.Cells(excel_row, date_col                          ).Value = work_assignment_date
        ObjExcel.Cells(excel_row, worker_id_col                     ).Value = qi_worker_id_number
        ObjExcel.Cells(excel_row, worker_name_col                   ).Value = qi_worker_full_name
        ObjExcel.Cells(excel_row, cases_reviewed_col                ).Value = exp_number_of_cases_reviewed
        ObjExcel.Cells(excel_row, cases_xfs_app_col                 ).Value = exp_number_of_cases_approved
        ObjExcel.Cells(excel_row, cases_xfs_no_id_col               ).Value = exp_number_of_cases_no_id
        ObjExcel.Cells(excel_row, cases_xfs_id_app                  ).Value = exp_number_of_cases_ID_approved
        ObjExcel.Cells(excel_row, cases_xfs_correct_col             ).Value = exp_number_of_cases_correct
        ObjExcel.Cells(excel_row, cases_xfs_no_caf                  ).Value = exp_number_of_cases_xfs_no_caf
        ObjExcel.Cells(excel_row, cases_xfs_verifs_not_postponed_col).Value = exp_number_of_cases_should_have_been_postponed
        ObjExcel.Cells(excel_row, cases_xfs_MAXIS_wrong_col         ).Value = exp_number_of_cases_MAXIS_incorrect
        ObjExcel.Cells(excel_row, cases_xfs_bad_note_col            ).Value = exp_number_of_cases_bad_notes
        ObjExcel.Cells(excel_row, xfs_assignment_length_col         ).Value = assignment_time
        ObjExcel.Cells(excel_row, xfs_assignment_assessment_col     ).Value = assignment_assesment
        ObjExcel.Cells(excel_row, xfs_list_of_cases_col             ).Value = assignment_case_numbers_to_save
		ObjExcel.Cells(excel_row, other_notes_col             		).Value = assignment_other_notes

        ObjExcel.ActiveWorkbook.Save        'saving and closing the Excel spreadsheet
        ObjExcel.ActiveWorkbook.Close
        ObjExcel.Application.Quit

		file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\QI Expedited Review " & date_for_doc & ".xlsx" 'single assignment file
		Set fso = CreateObject("Scripting.FileSystemObject")

		If (fso.FileExists(file_selection_path)) Then
			Call excel_open(file_selection_path, True, True, ObjWorkExcel, objWorkbook)  'opens the selected excel file'

			excel_row = 2
			Do
				For excel_col = 9 to 18
					If ObjWorkExcel.Cells(excel_row,  excel_col).Value = "Yes" Then ObjWorkExcel.Cells(excel_row,  excel_col).Value = "X"
				Next
				excel_row = excel_row + 1
				the_case_number = trim(ObjWorkExcel.Cells(excel_row, 2).Value)
			Loop Until the_case_number = ""
			'Saves and closes the most recent Excel workbook
			ObjWorkExcel.ActiveWorkbook.Save
			ObjWorkExcel.ActiveWorkbook.Close
			ObjWorkExcel.Application.Quit
			ObjWorkExcel.Quit
		End If

		MsgBox "Your counts and details have been successfully saved. The script will now send the emails to report out that your work assignment is completed."

        'Setting the beginning of the emails and the subject lines.
        main_email_body = "The Expedited Work Assignment has been completed for " & work_assignment_date & "."
        main_email_subject = "Expedited Work Assignment Completed"
        If assignment_case_numbers_to_save <> "" Then
            case_numbers_email_body = "Please add these cases to our next QI meeting agenda under the 'Exp SNAP Check-in'."
            case_numbers_email_subject = "EXP SNAP Assignment - Case Numbers to review"
        End If
        If assignment_new_ideas <> "" Then
            ideas_email_body = "New ideas for counts to do on Expedited processing. While processing the Expedited work assignment, I have noticed something that may be a trend we want to track. Please review looking at adding a count option for:"
            ideas_email_subject = "EXP SNAP Assignment - new ideas for counts and statistics"
        End If

    Case "Pending at Day 30 - Part of On Demand"                                'This opetion is for the Day 30 assignment
        counts_number = 6                                                       'There are 7 things we count so the array goes up to 6
        'Setting the file locations and doc strings
        word_doc_name = qi_worker_full_name & " - Day 30 Assignment Report for " & date_for_doc
        word_doc_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Applications Statistics\Day Thirty Assignments\" & month_of_assignment &"-" & year_of_assignment & "\"
        excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Applications Statistics\DAY THIRTY Work Counts.xlsx"

        Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 586, 235, "Details of Cases Pending over 30 Days Work Assignment"
		  EditBox 80, 40, 30, 15, d30_number_of_cases_reviewed
		  EditBox 125, 60, 30, 15, d30_number_of_cases_denied_no_interview
		  EditBox 130, 80, 30, 15, d30_number_of_cases_denied_other
		  EditBox 80, 100, 30, 15, d30_number_of_cases_approved
		  EditBox 105, 120, 30, 15, d30_number_of_cases_timely
		  EditBox 115, 140, 30, 15, d30_number_of_cases_not_timely
		  EditBox 160, 160, 30, 15, d30_number_of_cases_future_verif
		  Text 80, 10, 420, 10, "********* Pending over 30 Days - the review of the daily assignment of cases pending over 30 days - part of On Demand *********"
		  GroupBox 10, 25, 245, 155, "Number of cases that:"
		  Text 20, 45, 60, 10, ". you reviewed:"
		  Text 20, 65, 100, 10, "... you denied for no interview:"
		  Text 20, 85, 105, 10, "... you denied for other reasons:"
		  Text 20, 105, 60, 10, "... you approved:"
		  Text 20, 125, 80, 10, "... were acted on timely:"
		  Text 20, 145, 90, 10, "... were acted on NOT timely:"
		  Text 20, 165, 140, 10, "... have a future due date for verifications:"
		  EditBox 405, 35, 20, 15, assignment_hours
		  EditBox 455, 35, 20, 15, assignment_minutes
		  DropListBox 400, 60, 110, 45, "Select One..."+chr(9)+"Great"+chr(9)+"Good"+chr(9)+"Okay"+chr(9)+"Neutral"+chr(9)+"A little rough"+chr(9)+"Bad"+chr(9)+"Terrible", assignment_assesment
		  EditBox 265, 95, 285, 15, assignment_case_numbers_to_save
		  EditBox 265, 135, 285, 15, assignment_new_ideas
		  EditBox 10, 195, 570, 15, assignment_other_notes
		  ButtonGroup ButtonPressed
		    OkButton 480, 215, 50, 15
		    CancelButton 530, 215, 50, 15
		  Text 265, 40, 140, 10, "About how long did the assignment take?"
		  Text 430, 40, 20, 10, "hours"
		  Text 480, 40, 30, 10, "minutes"
		  Text 265, 65, 135, 10, "How was the assignment for you today?"
		  Text 265, 85, 180, 10, "Any case numbers to save for example/larger reivew?"
		  Text 265, 125, 105, 10, "Ideas of other counts to collect:"
		  Text 265, 150, 285, 15, "These are common erros or handling that we are seeing in review, this would be to add to the option on the left."
		  Text 10, 185, 140, 10, "Other notes about assignment from today:"
		EndDialog

        Do
            Do
                err_msg = ""

                dialog Dialog1
                cancel_confirmation

                If IsNumeric(d30_number_of_cases_reviewed) = FALSE OR IsNumeric(d30_number_of_cases_denied_no_interview) = FALSE OR IsNumeric(d30_number_of_cases_denied_other) = FALSE OR IsNumeric(d30_number_of_cases_approved) = FALSE OR IsNumeric(d30_number_of_cases_timely) = FALSE OR IsNumeric(d30_number_of_cases_not_timely) = FALSE OR IsNumeric(d30_number_of_cases_future_verif) = FALSE Then
                    err_msg = err_msg & vbNewLine & "* Count needed. Enter the number of cases that meet the following criteria: "
                    If IsNumeric(d30_number_of_cases_reviewed) = FALSE Then err_msg = err_msg & vbNewLine & "  - total you reviewed (1st)"
                    If IsNumeric(d30_number_of_cases_denied_no_interview) = FALSE Then err_msg = err_msg & vbNewLine & "  - you denied for no interview (2nd)"
                    If IsNumeric(d30_number_of_cases_denied_other) = FALSE Then err_msg = err_msg & vbNewLine & "  - you denied for other reasons (3rd)"
                    If IsNumeric(d30_number_of_cases_approved) = FALSE Then err_msg = err_msg & vbNewLine & "  - you approved (4th)"
                    If IsNumeric(d30_number_of_cases_timely) = FALSE Then err_msg = err_msg & vbNewLine & "  - were processed timely (5th)"
                    If IsNumeric(d30_number_of_cases_not_timely) = FALSE Then err_msg = err_msg & vbNewLine & "  - were processed, but not timely (6th)"
                    If IsNumeric(d30_number_of_cases_future_verif) = FALSE Then err_msg = err_msg & vbNewLine & "  - have a future date for verifications due (7th)"
                    err_msg = err_msg & vbNewLine & "  These counts are crucial to tracking our progress and advancement in our work assignment efforts." & vbNewLine
                End If

                If IsNumeric(assignment_hours) = FALSE AND IsNumeric(assignment_minutes) = FALSE Then err_msg = err_msg & vbNewLine & "* Enter how long it took for you to complete the assignment. This can be entered in hours and/or minutes. Do your best to discount breaks and other work so we get a good idea of how long the work is taking."
                If assignment_assesment = "Select One..." Then err_msg = err_msg & vbNewLine & "* Let us know how the work was for you today."

                If err_msg <> "" Then MsgBox "Please resolve to continue:" & vbNewLine & err_msg

            Loop until err_msg = ""
            Call check_for_password(are_we_passworded_out)
        Loop until are_we_passworded_out = FALSE

		leave_alone_msg = MsgBox("Thank you for providing all the detail." & vbNewLine & vbNewLine & "The script now needs to save all of the information, this can take a minute or two. The script needs to open Excel, Word, and send emails. This will work best if the script does not get interrupted." & vbNewLine & vbNewLine & "Do not touch the computer until the script finishes.", vbExclamation + vbSystemModal, "Let the Script Work Uninterrupted")

        If IsNumeric(assignment_minutes) = FALSE Then assignment_minutes = 0    'formatting the date
        If IsNumeric(assignment_hours) = TRUE Then                              'finding the time in minutes only for the spreadsheet
            minutes_from_hours = assignment_hours * 60
            assignment_time = assignment_minutes + minutes_from_hours
        Else
            assignment_time = assignment_minutes
        End If

        call excel_open(excel_file_path, False,  False, ObjExcel, objWorkbook)  'opening the Excel'

        sheet_selection = month_of_assignment &"-" & year_of_assignment         'Using the month and year to open the correct spreadsheet, which is named MM-YYYY
        ObjExcel.worksheets(sheet_selection).Activate

        excel_row = 2                                                           'finding the first empty excel row
        Do
            this_entry = ObjExcel.Cells(excel_row, 1).Value
            this_entry = trim(this_entry)
            If this_entry <> "" Then excel_row = excel_row + 1
        Loop until this_entry = ""

        'Adding the information from the dialog into the Excel spreadsheet
        ObjExcel.Cells(excel_row, date_col                      ).Value = work_assignment_date
        ObjExcel.Cells(excel_row, worker_id_col                 ).Value = qi_worker_id_number
        ObjExcel.Cells(excel_row, worker_name_col               ).Value = qi_worker_full_name
        ObjExcel.Cells(excel_row, cases_reviewed_col            ).Value = d30_number_of_cases_reviewed
        ObjExcel.Cells(excel_row, cases_d30_no_interview        ).Value = d30_number_of_cases_denied_no_interview
        ObjExcel.Cells(excel_row, cases_d30_other_reason        ).Value = d30_number_of_cases_denied_other
        ObjExcel.Cells(excel_row, cases_d30_app                 ).Value = d30_number_of_cases_approved
        ObjExcel.Cells(excel_row, cases_d30_timely              ).Value = d30_number_of_cases_timely
        ObjExcel.Cells(excel_row, cases_d30_not_timely          ).Value = d30_number_of_cases_not_timely
        ObjExcel.Cells(excel_row, cases_d30_future_verifs       ).Value = d30_number_of_cases_future_verif
        ObjExcel.Cells(excel_row, d30_assignment_length_col     ).Value = assignment_time
        ObjExcel.Cells(excel_row, d30_assignment_assessment_col ).Value = assignment_assesment
        ObjExcel.Cells(excel_row, d30_list_of_cases_col         ).Value = assignment_case_numbers_to_save

        ObjExcel.ActiveWorkbook.Save                                            'saving and closing the Excel spreadsheet
        ObjExcel.ActiveWorkbook.Close
        ObjExcel.Application.Quit

        ReDim TABLE_ARRAY(1, counts_number)                                     'sizing the array for this work assignment type

        TABLE_ARRAY(1, 0) = d30_number_of_cases_reviewed                        'saving the information to the array
        TABLE_ARRAY(1, 1) = d30_number_of_cases_denied_no_interview
        TABLE_ARRAY(1, 2) = d30_number_of_cases_denied_other
        TABLE_ARRAY(1, 3) = d30_number_of_cases_approved
        TABLE_ARRAY(1, 4) = d30_number_of_cases_timely
        TABLE_ARRAY(1, 5) = d30_number_of_cases_not_timely
        TABLE_ARRAY(1, 6) = d30_number_of_cases_future_verif

        TABLE_ARRAY(0, 0) = "Cases Reviewed"
        TABLE_ARRAY(0, 1) = "Denied - No Interview"
        TABLE_ARRAY(0, 2) = "Denied - Other Reason"
        TABLE_ARRAY(0, 3) = "Cases Approved"
        TABLE_ARRAY(0, 4) = "Cases Processed Timely"
        TABLE_ARRAY(0, 5) = "Cases NOT Timely"
        TABLE_ARRAY(0, 6) = "Cases with Future Verif Date"

        'Setting the first line and subject for the emails.
        main_email_body = "The Day Thirty Assignment has been completed for " & work_assignment_date & "."
        main_email_subject = "Day Thirty Assignment Completed"
        If assignment_case_numbers_to_save <> "" Then
            case_numbers_email_body = "Please add these cases to our next QI meeting agenda under the 'On Demand Check-in'."
            case_numbers_email_subject = "DAY THIRTY Assignment - Case Numbers to review"
        End If
        If assignment_new_ideas <> "" Then
            ideas_email_body = "New ideas for counts to do on Day Thirty processing. While processing the Day Thirty work assignment, I have noticed something that may be a trend we want to track. Please review looking at adding a count option for:"
            ideas_email_subject = "DAY THIRTY Assignment - new ideas for counts and statistics"
        End If
End Select

'Adding the rest of the detail to the email body
main_email_body = main_email_body & vbCr & "Completed by: " & qi_worker_full_name
main_email_body = main_email_body & vbCr & vbCr & "Review completed, information about the work completed today can be found at:"
' main_email_body = main_email_body & vbCr & "<" & word_doc_file_path & word_doc_name & ".docx>"
' main_email_body = main_email_body & vbCr & vbCr & "Count Worksheet updated, can be found: "
main_email_body = main_email_body & vbCr & "<" & today_file_selection_path & ">"
main_email_body = main_email_body & vbCr & "This will work for " & date & ". If it is after " & date & " the file can be found at:"
main_email_body = main_email_body & vbCr & "<" & archive_file_selection_path & ">"
main_email_body = main_email_body & vbCr & vbCr & "Work assignment assesment: " & assignment_assesment
main_email_body = main_email_body & vbCr & vbCr & "Length of assignment: " & assignment_hours & " hours and " & assignment_minutes & " minutes."

assignment_case_numbers_to_save = trim(assignment_case_numbers_to_save)
assignment_new_ideas = trim(assignment_new_ideas)
If assignment_case_numbers_to_save <> "" Then
    main_email_body = main_email_body & vbCr & vbCr & "Case numbers to discuss sent to QI email to be added to meeting agenda. Case numbers:"
    main_email_body = main_email_body & vbCr & assignment_case_numbers_to_save

    case_numbers_email_body = case_numbers_email_body & vbCr & assignment_case_numbers_to_save
    case_numbers_email_body = case_numbers_email_body & vbCr & vbCR & "These cases should be reviewed by the whole QI team and follow up decisions made."
    case_numbers_email_body = case_numbers_email_body & vbCr & vbCr & "------"
    case_numbers_email_body = case_numbers_email_body & vbCr & email_signature
    STATS_manualtime = STATS_manualtime + 120
End If
If assignment_new_ideas <> "" Then
    main_email_body = main_email_body & vbCr & vbCr & "New ideas for statistics to gather sent to the BZST. Ideas:"
    main_email_body = main_email_body & vbCr & assignment_new_ideas

    ideas_email_body = ideas_email_body & vbCr & assignment_new_ideas
    ideas_email_body = ideas_email_body & vbCr & vbCr & "------"
    ideas_email_body = ideas_email_body & vbCr & email_signature
    STATS_manualtime = STATS_manualtime + 120
End If
main_email_body = main_email_body & vbCr & vbCr & "------"
main_email_body = main_email_body & vbCr & email_signature


'send the list of messy cases to the QI email that Mandora reviews
' If assignment_case_numbers_to_save <> "" Then CALL create_outlook_email("HSPH.EWS.QUALITYIMPROVEMENT@hennepin.us", "", case_numbers_email_subject, case_numbers_email_body, "", TRUE)
'send the new ideas of things to count to the BZST email
' If assignment_new_ideas <> "" Then CALL create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", ideas_email_subject, ideas_email_body, "", TRUE)
'email all the people that this is done
' CALL create_outlook_email(qi_worker_supervisor_email, "Ilse.Ferris@hennepin.us", main_email_subject, main_email_body, "", TRUE)
CALL create_outlook_email("HSPH.EWS.BlueZoneScripts@hennepin.us", "", main_email_subject, main_email_body, "", TRUE)

Call script_end_procedure_with_error_report("Great work! Thank you for completing your assignment report.")
