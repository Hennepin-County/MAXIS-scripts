'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - REVIEW QCR REPORTS.vbs"
start_time = timer
STATS_counter = 0               'sets the stats counter at one
STATS_manualtime = 60          'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("09/19/2024", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------
'This script does not access MAXIS and so there is no EMConnect or background checks

'Define the dialog to select which kind of reviews to select.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 446, 110, "Select QCR Reports to Review"
  CheckBox 20, 25, 95, 10, "SNAP - ABAWD 30/09", snap_abawd_30_09_checkbox
  CheckBox 20, 40, 130, 10, "SNAP - Homeless Shelter Expense", snap_homeless_shelter_checkbox
  CheckBox 20, 55, 130, 10, "SNAP - UHFS Shelter Expense", snap_uhfs_shelter_checkbox
  CheckBox 165, 25, 130, 10, "HC - Remedial Care Deduction", hc_remedial_care_checkbox
  CheckBox 165, 40, 130, 10, "MSA - Shelter Needy", msa_shelter_needy_checkbox
  CheckBox 165, 55, 130, 10, "MSA - UNEA 44", msa_unea_44_checkbox
  EditBox 150, 70, 50, 15, date_cutoff
  DropListBox 125, 90, 165, 45, "Yes - Archive all Recorded Files"+chr(9)+"No - Leave the Files in Place", archive_files
  ButtonGroup ButtonPressed
    CancelButton 385, 90, 50, 15
    OkButton 330, 90, 50, 15
	PushButton 335, 15, 75, 15, "INSTRUCTIONS", instructions_btn
  Text 10, 10, 135, 10, "Select the type of QCR Reviews to Pull."
  Text 150, 10, 85, 10, "Check all that apply:"
  Text 10, 95, 115, 10, "Move Files to the Archive Folder?"
  Text 10, 75, 140, 10, "Select only reports BEFORE specific date:"
  Text 205, 75, 110, 10, "(Leave blank to pull all reports.)"
  Text 335, 5, 100, 10, "ADMIN - Review QCR Reports"
EndDialog

Do
	dialog Dialog1
	cancel_without_confirmation
	err_msg = ""

	'Ensure a selection is made - so we pull at least one type of QCR
	one_selection_made = False
	If snap_abawd_30_09_checkbox = checked Then one_selection_made = True
	If snap_homeless_shelter_checkbox = checked Then one_selection_made = True
	If snap_uhfs_shelter_checkbox = checked Then one_selection_made = True
	If hc_remedial_care_checkbox = checked Then one_selection_made = True
	If msa_shelter_needy_checkbox = checked Then one_selection_made = True
	If msa_unea_44_checkbox = checked Then one_selection_made = True

	'error message handling for the dialog
	If one_selection_made = False Then err_msg = err_msg & vbCr & "* Select at least one QCR type to continue."
	If date_cutoff <> "" Then
		If IsDate(date_cutoff) = False Then err_msg = err_msg & vbCr & "* The date cutoff does not appear to be a valid date, check and reenter."
	End If

	If ButtonPressed = instructions_btn Then
		run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ADMIN/ADMIN%20-%20REVIEW%20QCR%20REPORTS.docx"
		err_msg = "LOOP"
	Else
		If err_msg <> "" Then MsgBox "* * * *  NOTICE * * * *" & vbCr & err_msg
	End If

Loop until err_msg = ""

'set the column numbers
worker_id_col 		= 1
worker_name_col 	= 2
case_number_col 	= 3
run_date_col 		= 4
run_time_col 		= 5
program_col 		= 6
initial_elig_mo_col = 7
qcr_type_col 		= 8
policy_col 			= 9

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Create the header row
ObjExcel.Cells(1, worker_id_col) = "Worker ID"
ObjExcel.Cells(1, worker_name_col) = "Worker Name"
ObjExcel.Cells(1, case_number_col) = "Case Number"
ObjExcel.Cells(1, run_date_col) = "Script Run Date"
ObjExcel.Cells(1, run_time_col) = "Script Run Time"
ObjExcel.Cells(1, program_col) = "Program"
ObjExcel.Cells(1, initial_elig_mo_col) = "Initial ELIG Month"
ObjExcel.Cells(1, qcr_type_col) = "QCR Type"
ObjExcel.Cells(1, policy_col) = "Policy"

'This part of the header row dependent on which type of QCR Reports were selected
col_to_use = 10
If snap_abawd_30_09_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "ABAWD MEMBs"
	abawd_membs_col = col_to_use
	col_to_use = col_to_use + 1
End If
If snap_homeless_shelter_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "Homeless"
	hmls_col = col_to_use
	col_to_use = col_to_use + 1

	ObjExcel.Cells(1, col_to_use) = "Budgeted SHEL Expense"
	hmls_shel_expense_col = col_to_use
	col_to_use = col_to_use + 1
End If
If snap_uhfs_shelter_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "MEMBs on UHFS"
	uhfs_membs_col = col_to_use
	col_to_use = col_to_use + 1

	ObjExcel.Cells(1, col_to_use) = "MEMBs with SHEL"
	shel_membs_col = col_to_use
	col_to_use = col_to_use + 1
End If
If hc_remedial_care_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "Input for Remedial Care"
	remedial_care_input_col = col_to_use
	col_to_use = col_to_use + 1
End If
If msa_shelter_needy_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "Shelter Needy Amount"
	shel_need_amt_col = col_to_use
	col_to_use = col_to_use + 1
End If
If msa_unea_44_checkbox = checked Then
	ObjExcel.Cells(1, col_to_use) = "MEMBs with UNEA 44"
	unea_44_membs_col = col_to_use
	col_to_use = col_to_use + 1
End If

FOR i = 1 to col_to_use-1		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT
excel_row = 2

'setting some defaults for the file information to read
FILES_TO_MOVE_ARRAY = ""
tally = 0

'Read through all of the files in the QCR Log folder
Set objFolder = objFSO.GetFolder(t_drive & "\Eligibility Support\Assignments\QCR Logs")										'Creates an oject of the whole my documents folder
Set colFiles = objFolder.Files																'Creates an array/collection of all the files in the folder
For Each objFile in colFiles																'looping through each file
	qcr_file_name = objFile.Name															'Grabing the file name, path, and creation date
	qcr_file_path = objFile.Path
	qcr_file_date = objFile.DateCreated

	'reset the variables for each loop
	qcr_type = ""
	qcr_case_number = ""
	qcr_worker_number = ""
	qcr_worker_name = ""
	qcr_run_date_time = ""
	qcr_elig_program = ""
	qcr_initial_elig_month = ""
	qcr_policy = ""
	qcr_abawd_membs = ""
	qcr_homeless = ""
	qcr_budgeted_shel = ""
	qcr_uhfs_membs = ""
	qcr_shel_membs = ""
	qcr_remedial_care_input = ""
	qcr_shelter_needy_amt = ""
	qcr_unea_44_membs = ""

	'determining which type of QCR Report the file selected is.
	If InStr(qcr_file_name, "SNAP_WREG_30_09_") <> 0 Then qcr_type = "SNAP - ABAWD 30 09"
	If InStr(qcr_file_name, "SNAP_review_homeless_shelter_expense_") <> 0 Then qcr_type = "SNAP - Homeless Shelter Expense"
	If InStr(qcr_file_name, "SNAP_review_homeless_shleter_expense_") <> 0 Then qcr_type = "SNAP - Homeless Shelter Expense"
	If InStr(qcr_file_name, "SNAP_UHFS_SHEL_Expense_") <> 0 Then qcr_type = "SNAP - UHFS Shelter Expense"
	If InStr(qcr_file_name, "HC_Remedial_Care_") <> 0 Then qcr_type = "HC - Remedial Care Deduction"
	If InStr(qcr_file_name, "MSA_Shelter_Needy_") <> 0 Then qcr_type = "MSA - Shelter Needy"
	If InStr(qcr_file_name, "MSA_UNEA_44_") <> 0 Then qcr_type = "MSA - UNEA 44"

	'Identify if the file should be read based on the QCR type
	read_file = False
	If snap_abawd_30_09_checkbox = checked and qcr_type = "SNAP - ABAWD 30 09" Then read_file = True
	If snap_homeless_shelter_checkbox = checked and qcr_type = "SNAP - Homeless Shelter Expense" Then read_file = True
	If snap_uhfs_shelter_checkbox = checked and qcr_type = "SNAP - UHFS Shelter Expense" Then read_file = True
	If hc_remedial_care_checkbox = checked and qcr_type = "HC - Remedial Care Deduction" Then read_file = True
	If msa_shelter_needy_checkbox = checked and qcr_type = "MSA - Shelter Needy" Then read_file = True
	If msa_unea_44_checkbox = checked and qcr_type = "MSA - UNEA 44" Then read_file = True

	'If a data limitation exists, determine if the file should be read based on the file date
	If IsDate(date_cutoff) = True Then
		If DateDiff("d", qcr_file_date, date_cutoff) =< 0 Then read_file = False
	End If

	'If the file is the right type and within the date range (if selected)
	If read_file = True Then
		With (CreateObject("Scripting.FileSystemObject"))
			'Creating an object for the stream of text which we'll use frequently
			Dim objTextStream

			If .FileExists(qcr_file_path) = True then
				'Setting the object to open the text file for reading the data already in the file
				Set objTextStream = .OpenTextFile(qcr_file_path, ForReading)
				FILES_TO_MOVE_ARRAY = FILES_TO_MOVE_ARRAY & "~!~" & qcr_file_path			'This is a list of cases to archive if this option is selected

				'Reading the entire text file into a string
				every_line_in_text_file = objTextStream.ReadAll

				'Splitting the text file contents into an array which will be sorted
				qcr_details = split(every_line_in_text_file, vbNewLine)

				'Read each line of the file.
				For Each text_line in qcr_details
					If InStr(text_line, "^&*^&*") <> 0 Then
						line_items_array = ""
						line_items_array = split(text_line, "^&*^&*")

						If line_items_array(0) = "WorkerNumber" Then qcr_worker_number = line_items_array(1)
						If line_items_array(0) = "WorkerName" Then qcr_worker_name = line_items_array(1)
						If line_items_array(0) = "RunDateTime" Then qcr_run_date_time = line_items_array(1)
						If line_items_array(0) = "Case Number" Then qcr_case_number = line_items_array(1)
						If line_items_array(0) = "ELIGProgram" Then qcr_elig_program = line_items_array(1)
						If line_items_array(0) = "InitialELIGMonthInPackage" Then
							qcr_initial_elig_month = line_items_array(1)
							qcr_initial_elig_month = replace(qcr_initial_elig_month, "/", "/1/")
						End If
						If line_items_array(0) = "POLICY" Then qcr_policy = line_items_array(1)
						If line_items_array(0) = "ABAWDMembs" Then qcr_abawd_membs = line_items_array(1)
						If line_items_array(0) = "Homeless" Then qcr_homeless = line_items_array(1)
						If line_items_array(0) = "BudgetedSHELExpense" Then qcr_budgeted_shel = line_items_array(1)
						If line_items_array(0) = "UHFSMembs" Then qcr_uhfs_membs = line_items_array(1)
						If line_items_array(0) = "SHELMembs" Then qcr_shel_membs = line_items_array(1)
						If line_items_array(0) = "HCBudgNoRemedialCareInput" Then qcr_remedial_care_input = line_items_array(1)
						If line_items_array(0) = "ShelterNeedyAmount" Then qcr_shelter_needy_amt = line_items_array(1)
						If line_items_array(0) = "UNEA44Membs" Then qcr_unea_44_membs = line_items_array(1)

					End If
				Next

				'Add each piece of the file information to the created Excel file
				ObjExcel.Cells(excel_row, worker_id_col) = qcr_worker_number
				ObjExcel.Cells(excel_row, worker_name_col) = qcr_worker_name
				ObjExcel.Cells(excel_row, case_number_col) = qcr_case_number
				ObjExcel.Cells(excel_row, run_date_col) = FormatDateTime(qcr_run_date_time, 2)
				ObjExcel.Cells(excel_row, run_time_col) = FormatDateTime(qcr_run_date_time, 3)
				ObjExcel.Cells(excel_row, program_col) = qcr_elig_program
				ObjExcel.Cells(excel_row, initial_elig_mo_col) = qcr_initial_elig_month
				ObjExcel.Cells(excel_row, qcr_type_col) = qcr_type
				ObjExcel.Cells(excel_row, policy_col) = qcr_policy

				If snap_abawd_30_09_checkbox = checked Then ObjExcel.Cells(excel_row, abawd_membs_col) = qcr_abawd_membs
				If snap_homeless_shelter_checkbox = checked Then
					ObjExcel.Cells(excel_row, hmls_col) = qcr_homeless
					ObjExcel.Cells(excel_row, hmls_shel_expense_col) = qcr_budgeted_shel
				End If
				If snap_uhfs_shelter_checkbox = checked Then
					ObjExcel.Cells(excel_row, uhfs_membs_col) = qcr_uhfs_membs
					ObjExcel.Cells(excel_row, shel_membs_col) = qcr_shel_membs
				End If
				If hc_remedial_care_checkbox = checked Then ObjExcel.Cells(excel_row, remedial_care_input_col) = qcr_remedial_care_input
				If msa_shelter_needy_checkbox = checked Then ObjExcel.Cells(excel_row, shel_need_amt_col) = qcr_shelter_needy_amt
				If msa_unea_44_checkbox = checked Then ObjExcel.Cells(excel_row, unea_44_membs_col) = qcr_unea_44_membs

				objTextStream.Close
				excel_row = excel_row + 1
				STATS_counter = STATS_counter + 1
			End If
		End With
	End If
Next

'Format the Excel to size the column width
For col_to_autofit = 1 to col_to_use-1
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'If it was selected to archive the files, this part will move each file to the Archive folder
If archive_files = "Yes - Archive all Recorded Files" Then
	STATS_manualtime = 75
	FILES_TO_MOVE_ARRAY = trim(FILES_TO_MOVE_ARRAY)
	If left(FILES_TO_MOVE_ARRAY, 3) = "~!~" Then FILES_TO_MOVE_ARRAY = right(FILES_TO_MOVE_ARRAY, len(FILES_TO_MOVE_ARRAY)-3)
	FILES_TO_MOVE_ARRAY = split(FILES_TO_MOVE_ARRAY, "~!~")
	txt_file_archive_path = t_drive & "\Eligibility Support\Assignments\QCR Logs\Archive"

	For each file_path in FILES_TO_MOVE_ARRAY
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set theFile = fso.GetFile(file_path)

		this_file_name = theFile.Name
		objFSO.MoveFile file_path , txt_file_archive_path & "\" & this_file_name & ".txt"    'moving each file to the archive file
	Next
End If

Call script_end_procedure("Excel File Created with Information from the QCR Reports.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/19/2024
'--Tab orders reviewed & confirmed----------------------------------------------09/19/2024
'--Mandatory fields all present & Reviewed--------------------------------------09/19/2024
'--All variables in dialog match mandatory fields-------------------------------09/19/2024
'Review dialog names for content and content fit in dialog----------------------09/19/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------
'--Create a button to reference instructions------------------------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------