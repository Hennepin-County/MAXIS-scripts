'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - Work Assignment from Excel.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 60			 'manual run time in seconds
STATS_denomination = "I"		 'C is for each case
'END OF stats block==============================================================================================
'run_locally = TRUE
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
call changelog_update("07/11/2018", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function excel_open_pw(file_url, visible_status, alerts_status, ObjExcel, objWorkbook, my_password)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
    objExcel.DisplayAlerts = alerts_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url,,,, my_password) 'Opens an excel file from a specific URL
    ''(file.Path,,,, "mypassword",,,,,,,,,,)

end function


'END FUNCTION BLOCK ========================================================================================================

'DECLARATIONS ==============================================================================================================
'The columns in the Master Excel List
Const case_number_col           = 1
Const pmi_number_col            = 2
Const member_ref_number_col     = 3
Const last_name_col             = 4
Const first_name_col            = 5
Const april_status_col          = 6
Const may_status_col            = 7
Const june_status_col           = 8
Const july_status_col           = 9
Const august_status_col         = 10
Const sept_status_col           = 11
Const banked_months_count_col   = 12
Const not_abawd_col             = 13
Const homeless_WCOM_col         = 14
Const currently_assigned_col    = 15
Const notes_col                 = 16

Const xl_col_1              = 1
Const xl_col_2              = 2
Const xl_col_3              = 3
Const xl_col_4              = 4
Const xl_col_5              = 5
Const xl_col_6              = 6
Const xl_col_7              = 7
Const xl_col_8              = 8
Const xl_col_9              = 9
Const xl_col_10             = 10
Const xl_col_11             = 11
Const xl_col_12             = 12
Const xl_col_13             = 13
Const xl_col_14             = 14
Const xl_col_15             = 15
Const xl_col_16             = 16
Const xl_col_17             = 17
Const xl_col_18             = 18
Const xl_col_19             = 19
Const xl_col_20             = 20
Const xl_col_21             = 21
Const xl_col_22             = 22
Const xl_col_23             = 23
Const xl_col_24             = 24
Const xl_col_25             = 25
Const row_already_assigned  = 26
Const master_xl_row         = 27
Const list_action           = 28

'constants for arrays
Const case_nbr              = 00
Const clt_excel_row         = 01
Const memb_ref_nbr          = 02
Const memb_pmi_nbr          = 03
Const clt_last_name         = 04
Const clt_first_name        = 05
Const clt_full_name         = 06
Const april_bm              = 07
Const may_bm                = 08
Const june_bm               = 09
Const july_bm               = 10
Const august_bm             = 11
Const sept_bm               = 12
Const not_abawd_boolean     = 13
Const homeless_wcom_exists  = 14
Const assigned_boolean      = 15
Const clt_notes             = 16

Const assign_case           = 17
Const array_notes           = 18

Dim MASTER_LIST_ALL_ROWS()
ReDim MASTER_LIST_ALL_ROWS(list_action, 0)

Const assigned_worker_name      = 1
Const assigned_worker_email     = 2
Const assignment_list_path      = 3
Const script_call_name          = 4
Const last_excel_row            = 5
Const new_assignment            = 6
Const list_message              = 7

Dim ASSIGNMENT_LISTS_ARRAY()
ReDim ASSIGNMENT_LISTS_ARRAY(list_message, 0)


Const column_header              = 1
Const include_column             = 2
Const master_list_col_letter     = 3
Const assignment_list_col_letter = 4
Const master_list_col_number     = 5
Const assignment_list_col_number = 6

Dim COLUMN_ARRAY()
ReDim COLUMN_ARRAY(assignment_list_col_number, 0)

'END DECLARATIONS BLOCK ====================================================================================================

'THE SCRIPT ================================================================================================================
'Connects to BlueZone
EMConnect ""

run_by_QI_leadership = False
If user_ID_for_validation = "TAPA002" then run_by_QI_leadership = True
If user_ID_for_validation = "ILFE001" then run_by_QI_leadership = True
If user_ID_for_validation = "WFS395" then run_by_QI_leadership = True
If user_ID_for_validation = "CALO001" then run_by_QI_leadership = True
If user_ID_for_validation = "WFX901" then run_by_QI_leadership = True
If user_ID_for_validation = "WFU851" then run_by_QI_leadership = True

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)

assignment_year = DatePart("yyyy", date)
assignment_year = assignment_year & ""
assignment_month = MonthName(DatePart("m", date))


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 386, 235, "Create Assignment Lists from a Master List"
  EditBox 95, 25, 235, 15, master_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 335, 25, 45, 15, "Browse...", select_a_file_button
  EditBox 180, 85, 175, 15, assignment_list_title
  DropListBox 165, 120, 60, 45, "Select One..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", assignment_month
  EditBox 235, 120, 35, 15, assignment_year
  EditBox 95, 150, 285, 15, assignment_project_detail
  If run_by_QI_leadership = True Then DropListBox 115, 180, 140, 45, "Select One..."+chr(9)+"Select from QI"+chr(9)+"Restart a Previous Run"+chr(9)+"Excel List"+chr(9)+"Manual Entry", worker_selection
  If run_by_QI_leadership = False Then DropListBox 115, 180, 140, 45, "Select One..."+chr(9)+"Restart a Previous Run"+chr(9)+"Excel List"+chr(9)+"Manual Entry", worker_selection
  CheckBox 10, 210, 260, 10, "Check here to have the script close the assignment lists at the end of the run.", close_assignment_lists_checkbox
  CheckBox 10, 220, 200, 10, "Check here to have the script send assignment emails.", send_email_checkbox
  ButtonGroup ButtonPressed
    OkButton 275, 215, 50, 15
    CancelButton 330, 215, 50, 15
  Text 10, 10, 365, 10, "This script will take a list of cases and create even assignment lists for a list of workers."
  Text 15, 30, 75, 10, "Select the Master List:"
  Text 95, 45, 280, 20, "The MASTER LIST must be an Excel document that has a list of cases to be assigned. The script will not filter or sort these cases at all, so be sure your list is accurate."
  Text 105, 65, 250, 15, "** Master list much include a column titled 'Assigned' to indicate if the case has been assigned by the script."
  Text 95, 90, 85, 10, "Title of Assignment Lists"
  Text 180, 105, 180, 10, "Will be added to the file names of the assignment list."
  Text 95, 125, 65, 10, "Assignment Month: "
  Text 95, 140, 95, 10, "Assignment Project Detail:"
  Text 95, 165, 155, 10, "(Information is for email of assignment only.)"
  Text 5, 185, 110, 10, "How will the workers be selected:"
  Text 120, 195, 125, 10, "Manual entry is limited to 15 workers."
EndDialog

Do
    Do
        err_msg = ""

        Dialog dialog1
        cancel_without_confirmation

        master_excel_file_path = trim(master_excel_file_path)
        assignment_list_title = trim(assignment_list_title)
        assignment_year = trim(assignment_year)

        If master_excel_file_path = "" Then err_msg = err_msg & vbNewLine & "* Select an Excel file of a list of items to be assigned."
        If right(master_excel_file_path, 5) <> ".xlsx" Then err_msg = err_msg & vbNewLine & "* The file selected does not appear to be an Excel File. Please review."
        If assignment_list_title = "" Then err_msg = err_msg & vbNewLine & "* Enter a short title for the type of assignment being created."
        If assignment_month = "Select One..." Then err_msg = err_msg & vbNewLine & "* Select the month of the assignment."
        If assignment_year = "" Then err_msg = err_msg & vbNewLine & "* Enter the year of the assignment."
        If worker_selection = "" Then err_msg = err_msg & vbNewLine & "* Select how you would like to select the workers to assign this work to."
		If send_email_checkbox = checked and trim(assignment_project_detail) = "" Then err_msg = err_msg & vbNewLine & "* Since you are sending Emails to the assignees, detail information about the assignments in the 'Project Assignment Detail' area."
		If ButtonPressed = select_a_file_button then
            call file_selection_system_dialog(master_excel_file_path, ".xlsx")
            err_msg = err_msg & "LOOP"
        Else
            If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue: " & vbNewLine & err_msg
        End If

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

call excel_open_pw(master_excel_file_path, True, False, ObjExcel, objWorkbook, "")
master_list_folder = ""
folder_breadcrumbs = split(master_excel_file_path, "\")
For each folder in folder_breadcrumbs
    If right(folder, 5) <> ".xlsx" then
		If folder = "T:" Then folder = t_drive
		master_list_folder = master_list_folder & folder & "\"
	End If
Next

xl_col = 1
col_array_counter = 0
work_list_columns = "Select One..."
Do
    col_header = trim(ObjExcel.Cells(1, xl_col).Value)
    If col_header <> "" Then
        ReDim Preserve COLUMN_ARRAY(assignment_list_col_number, col_array_counter)
        COLUMN_ARRAY(master_list_col_number, col_array_counter) = xl_col
        COLUMN_ARRAY(master_list_col_letter, col_array_counter) = convert_digit_to_excel_column(xl_col)
        COLUMN_ARRAY(column_header, col_array_counter) = col_header
        COLUMN_ARRAY(include_column, col_array_counter) = checked

        work_list_columns = work_list_columns & chr(9) & COLUMN_ARRAY(column_header, col_array_counter)
        If UCase(col_header) = "ASSIGNED" Then
            COLUMN_ARRAY(include_column, col_array_counter) = unchecked
            assigned_selection = COLUMN_ARRAY(column_header, col_array_counter)
            assigned_marked_col = xl_col
        End If
        xl_col = xl_col + 1
        col_array_counter = col_array_counter + 1
    End If
Loop until col_header = ""

dlg_hgt = 115 + UBound(COLUMN_ARRAY,2)*10
y_pos = 30
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, dlg_hgt, "List of Workers to Assign"
  Text 10, 10, 365, 10, "Check each column from the master list you want to include on the assignment list"
  For each_column = 0 to UBound(COLUMN_ARRAY,2)
      CheckBox 15, y_pos, 225, 10, "Column " & COLUMN_ARRAY(master_list_col_letter, each_column) & "  -  " & COLUMN_ARRAY(column_header, each_column), COLUMN_ARRAY(include_column, each_column)
      y_pos = y_pos + 10
  Next
  y_pos = y_pos + 5
  Text 10, y_pos, 365, 10, "Pick the column you want to use to track that the row has been assigned"
  y_pos = y_pos + 10
  Text 10, y_pos + 5, 70, 10, "Assignment Column:"
  DropListBox 80, y_pos, 75, 15, work_list_columns, assigned_selection
  y_pos = y_pos + 20
  Text 10, y_pos, 365, 10, "If you have a column to track who it was assigned to, select it here:"
  y_pos = y_pos + 10
  DropListBox 10, y_pos, 75, 15, work_list_columns, assigned_to_worker_selection
  ButtonGroup ButtonPressed
    OkButton 225, y_pos, 50, 15
EndDialog

Do
    Do
        dialog dialog1
        cancel_without_confirmation

        If assigned_selection = "Select One..." Then MsgBox "Script Cannot Continue" & vbNewLine & "*** NOTICE ***" & vbNewLine & vbNewLine & "A column for if the case has been assigned must be identified for the script to track progress. If there is no column for this assignment tracking, the script run should be cancelled and a column added and thes cript rerun."
    Loop until assigned_selection <> "Select One..."
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

col_nbr = 1
For each_column = 0 to UBound(COLUMN_ARRAY,2)
    ' MsgBox "Array header - " & COLUMN_ARRAY(column_header, each_column) & vbNewLine & "Assigned Selection - " & assigned_selection & "Col - " &
    If COLUMN_ARRAY(column_header, each_column) = assigned_selection Then
        assigned_marked_col = COLUMN_ARRAY(master_list_col_number, each_column)
    End If
	If COLUMN_ARRAY(column_header, each_column) = assigned_to_worker_selection Then
		assigned_worker_col = COLUMN_ARRAY(master_list_col_number, each_column)
	End If
    If COLUMN_ARRAY(include_column, each_column) = checked Then

    End If
Next

assignment_folder = master_list_folder

Dim objTextStream
assignment_status_path = assignment_folder & "assignment-status.txt"

If worker_selection = "Restart a Previous Run" Then

    With objFSO
        'Creating an object for the stream of text which we'll use frequently
        If .FileExists(assignment_status_path) = True then
            'Setting the object to open the text file for reading the data already in the file
            Set objTextStream = .OpenTextFile(assignment_status_path, ForReading)

            'Reading the entire text file into a string
            every_line_in_text_file = objTextStream.ReadAll

            'Splitting the text file contents into an array which will be sorted
            assignment_paths_array = split(every_line_in_text_file, vbNewLine)

            lists_counter = 0
            for i = 0 to ubound(assignment_paths_array)
                If assignment_paths_array(i) <> "" then 'some are likely blank
                    ReDim Preserve ASSIGNMENT_LISTS_ARRAY(list_message, list_counter)
                    split_place = InStr(assignment_paths_array(i), "~")

                    ASSIGNMENT_LISTS_ARRAY(assignment_list_path, list_counter) = trim(left(assignment_paths_array(i), split_place-1))
                    ASSIGNMENT_LISTS_ARRAY(script_call_name, list_counter) = trim(right(assignment_paths_array(i), len(assignment_paths_array(i))- split_place))

                    list_counter = list_counter + 1
                End If
            next
        Else
            call script_end_procedure("No file of a previous run can be found and the script cannot continue with the selected master list and restart option. Check your options and restart the script run.")
        End If
    End With

    If ASSIGNMENT_LISTS_ARRAY(assignment_list_path, 0) <> "" Then
        For each_list = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
            assignment_files_known = TRUE
            excel_name = ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list)
            ' MsgBox "File Name - " & ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_list) & vbNewLine & "Call - " & ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list)
            Call excel_open(assignment_folder & "\" & ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_list), TRUE, FALSE, ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list), objWorkbook)
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list).worksheets("Information").Activate
            ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_list) = ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list).Cells(1, 2). Value
            ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_list) = ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list).Cells(2, 2). Value

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list).worksheets("Assignment").Activate
            xl_row = 1
            Do
                xl_row = xl_row + 1
                the_case_number = ASSIGNMENT_LISTS_ARRAY(script_call_name, each_list).Cells(xl_row, 1).Value
                the_case_number = trim(the_case_number)
            Loop Until the_case_number = ""
            ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_list) = xl_row
            ' MsgBox ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_list)
        Next
    End If
    number_of_assignment_lists = UBound(ASSIGNMENT_LISTS_ARRAY,2)+ 1 & ""
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 151, 90, "Number of Lists"
      EditBox 120, 45, 25, 15, number_of_assignment_lists
      ButtonGroup ButtonPressed
        OkButton 95, 70, 50, 15
      Text 10, 10, 125, 25, UBound(ASSIGNMENT_LISTS_ARRAY,2)+ 1 & " Asignment lists were found. You can increase this number to add more. Additional lists will be manual entry."
      Text 10, 50, 100, 10, "How many assignment lists?"
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

    dlg_hgt = 90
    number_of_assignment_lists = number_of_assignment_lists * 1
    If number_of_assignment_lists <> UBound(ASSIGNMENT_LISTS_ARRAY, 2) + 1 Then
        ReDim Preserve ASSIGNMENT_LISTS_ARRAY(list_message, number_of_assignment_lists-1)
        add_new = TRUE
    Else
        dlg_hgt = dlg_hgt - 15
    End If

    count_to_three = 1
    For each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY,2)
        If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) <> "" Then
            If count_to_three = 1 Then dlg_hgt = dlg_hgt + 15
            count_to_three = count_to_three + 1
            If count_to_three = 4 then count_to_three = 1
        Else
            dlg_hgt = dlg_hgt + 20
        End If
    Next
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 640, dlg_hgt, "Assignment Details"
      Text 10, 10, 440, 10, "Enter the details about the assignments to be created here:"
      Text 15, 30, 250, 10, "Previously Created Assignment Lists"
      y_pos = 45
      x_pos = 15
      for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
          If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) <> "" Then
              Text x_pos, y_pos, 200, 10, ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) & " - " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker)
              ASSIGNMENT_LISTS_ARRAY(new_assignment, each_worker) = FALSE
              x_pos = x_pos + 205
              If x_pos = 630 Then
                  y_pos = y_pos + 15
                  x_pos = 15
              End If
          End If
      Next
      y_pos = y_pos + 5
      If add_new = TRUE Then
          Text 15, y_pos, 50, 10, "Worker Name"
          Text 135, y_pos, 50, 10, "Worker Email"
          y_pos = y_pos + 15
          for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
              If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) = "" Then
                  EditBox 15, y_pos, 100, 15, ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
                  EditBox 135, y_pos, 100, 15, ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker)
                  Text 240, y_pos + 5, 50, 10, "@hennepin.us"
                  ASSIGNMENT_LISTS_ARRAY(new_assignment, each_worker) = TRUE
                  new_lists_needed = TRUE
                  y_pos = y_pos + 20
              End If
          Next
      End If
      y_pos = y_pos + 5
      ButtonGroup ButtonPressed
        OkButton 530, y_pos, 50, 15
        CancelButton 585, y_pos, 50, 15
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation

            for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
                If ASSIGNMENT_LISTS_ARRAY(new_assignment, each_worker) = TRUE Then
                    ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) = trim(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker))
                    If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) = "" Then err_msg = err_msg & vbNewLine & "* You must list all workers."
                    If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker), " ") = 0 Then err_msg = err_msg & vbNewLine & "* Please enter a first and last name for each worker."

                    ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) = trim(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker))
                    If ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) = "" Then err_msg = err_msg & vbNewLine & "* Enter a worker email address for each person."
                    If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker), " ") <> 0 Then err_msg = err_msg & vbNewLine & "* There are spaces in the email address, please review. For " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
                    If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker), ".") = 0 Then err_msg = err_msg & vbNewLine & "* The email address seems to be incorrect, please review. for " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
                End If
            next

            If err_msg <> "" Then MsgBox "Please resolve to continue: " & vbNewLine & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

ElseIf worker_selection = "Select from QI" Then
	new_lists_needed = TRUE
	' If IsArray(tester_array) = False Then
		Dim tester_array()
		ReDim tester_array(0)

		tester_list_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\COMPLETE LIST OF TESTERS.vbs"        'Opening the list of testers - which is saved locally for security
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(tester_list_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	' End If

	const qi_worker_name_const 		= 0
	const qi_worker_email_const		= 1
	const qi_worker_checkbox_const	= 2
	const qi_worker_first_name_const= 3
	const qi_worker_last_const		= 10

	Dim QI_WORKERS_ARRAY()
	ReDim QI_WORKERS_ARRAY(qi_worker_last_const, 0)

	qi_worker_count = 0
	' MsgBox "Here we go"
	For tester = 0 to UBound(tester_array)                         'looping through all of the testers
		' MsgBox "tester - " & tester & vbCr & "tester_array(tester).tester_supervisor_name - " & tester_array(tester).tester_supervisor_name
		If tester_array(tester).tester_supervisor_name = "Tanya Payne" Then
			RedIm preserve QI_WORKERS_ARRAY(qi_worker_last_const, qi_worker_count)

			QI_WORKERS_ARRAY(qi_worker_name_const, qi_worker_count) = tester_array(tester).tester_full_name
			QI_WORKERS_ARRAY(qi_worker_first_name_const, qi_worker_count) = tester_array(tester).tester_first_name
			QI_WORKERS_ARRAY(qi_worker_email_const, qi_worker_count) = tester_array(tester).tester_email

			qi_worker_count = qi_worker_count + 1
		End If
	Next
	' MsgBox "qi_worker_count - " & qi_worker_count
	dlg_len = 65 + qi_worker_count*10

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 416, dlg_len, "QI Worker Selection"
	  Text 10, 5, 210, 10, "Select any of the QI Workers that you want to assign a list to."
	  Text 10, 20, 50, 10, "Name"
	  Text 120, 20, 50, 10, "Email"
	  Text 265, 20, 130, 10, "Check all that you want to assign to"
	  y_pos = 35
	  For each_worker = 0 to UBound(QI_WORKERS_ARRAY, 2)
		  Text 10, y_pos, 100, 10, QI_WORKERS_ARRAY(qi_worker_name_const, each_worker)
		  Text 120, y_pos, 145, 10, QI_WORKERS_ARRAY(qi_worker_email_const, each_worker)
		  CheckBox 265, y_pos, 125, 10, "Assign to " & QI_WORKERS_ARRAY(qi_worker_first_name_const, each_worker), QI_WORKERS_ARRAY(qi_worker_checkbox_const, each_worker)
		  y_pos = y_pos + 10
	  Next
	  y_pos = y_pos + 10
	  ' Text 10, 35, 100, 10, "Faughn Ramisch-Church"
	  ' Text 120, 35, 145, 10, "faughn.ramisch-church@hennepin.us"
	  ' CheckBox 265, 35, 125, 10, "Assign to WORKER", checkbox
	  ' Text 10, 45, 100, 10, "Faughn Ramisch-Church"
	  ' Text 120, 45, 145, 10, "faughn.ramisch-church@hennepin.us"
	  ' CheckBox 265, 45, 125, 10, "Assign to WORKER", Check2
	  ' Text 10, 55, 100, 10, "Faughn Ramisch-Church"
	  ' Text 120, 55, 145, 10, "faughn.ramisch-church@hennepin.us"
	  ' CheckBox 265, 55, 125, 10, "Assign to WORKER", Check3
	  ' Text 10, 65, 100, 10, "Faughn Ramisch-Church"
	  ' Text 120, 65, 145, 10, "faughn.ramisch-church@hennepin.us"
	  ' CheckBox 265, 65, 125, 10, "Assign to WORKER", Check4
	  ButtonGroup ButtonPressed
	    OkButton 305, y_pos, 50, 15
	    CancelButton 360, y_pos, 50, 15
	EndDialog

	Dialog Dialog1


	assigned_workers = 0

	For each_worker = 0 to UBound(QI_WORKERS_ARRAY, 2)
		If QI_WORKERS_ARRAY(qi_worker_checkbox_const, each_worker) = checked Then
			ReDim preserve ASSIGNMENT_LISTS_ARRAY(list_message, assigned_workers)

			 ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, assigned_workers) = QI_WORKERS_ARRAY(qi_worker_name_const, each_worker)
			 ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, assigned_workers) = QI_WORKERS_ARRAY(qi_worker_email_const, each_worker)
			 ASSIGNMENT_LISTS_ARRAY(new_assignment, assigned_workers) = True
			assigned_workers = assigned_workers + 1
		End If
	Next
	' MsgBox "Wait Here"

ElseIf worker_selection = "Excel List" Then
    new_lists_needed = TRUE
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 386, 80, "List of Workers to Assign"
      EditBox 95, 25, 235, 15, workers_list_excel_path
      ButtonGroup ButtonPressed
        PushButton 335, 25, 45, 15, "Browse...", select_a_file_button
        OkButton 275, 60, 50, 15
        CancelButton 330, 60, 50, 15
      Text 10, 10, 365, 10, "Select the existing Excel document with a list of workers to get an assignment"
      Text 15, 30, 75, 10, "Select the Worker List:"
      Text 95, 45, 280, 10, "The worker list should have a column with the worker name (first and last) and email."
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation

            If ButtonPressed = select_a_file_button then
                call file_selection_system_dialog(workers_list_excel_path, ".xlsx")
                err_msg = err_msg & "LOOP"
            Else
                If err_msg <> "" Then MsgBox "*** NOTICE ***" & vbNewLine & "Please resolve to continue: " & vbNewLine & err_msg
            End If

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    call excel_open_pw(workers_list_excel_path, True, False, ObjWorkersExcel, objWorkbook, "")

    const wrk_col_numb          = 0
    const wrk_col_head          = 1
    const wrk_col_head_format   = 2
    const wrk_col_letter        = 3
    const wrk_col_detail        = 4

    Dim WORKERS_COL_ARRAY()
    ReDim WORKERS_COL_ARRAY(wrk_col_detail, 0)
    worker_full_name_col = 0
    worker_first_name_col = 0
    worker_last_name_col = 0
    worker_email_col = 0
    worker_columns_list = "Select One..."

    excel_col = 1
    Do
        col_name = trim(ObjWorkersExcel.Cells(1, excel_col).Value)
        If col_name <> "" Then
            ReDim Preserve WORKERS_COL_ARRAY(wrk_col_detail, excel_col-1)
            WORKERS_COL_ARRAY(wrk_col_numb, excel_col-1) = excel_col
            WORKERS_COL_ARRAY(wrk_col_head, excel_col-1) = col_name
            WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) = convert_digit_to_excel_column(excel_col)
            worker_columns_list = worker_columns_list & chr(9) & "COLUMN " & WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) & " - " & WORKERS_COL_ARRAY(wrk_col_head, excel_col-1)

            formatted_col_name = ucase(col_name)
            formatted_col_name = replace(formatted_col_name, " ", "")
            WORKERS_COL_ARRAY(wrk_col_head_format, excel_col-1) = formatted_col_name
            If formatted_col_name = "LAST" then worker_last_name_col = excel_col
            If formatted_col_name = "NAME-LAST" then worker_last_name_col = excel_col
            If formatted_col_name = "LASTNAME" then worker_last_name_col = excel_col
            If formatted_col_name = "FIRST" then worker_first_name_col = excel_col
            If formatted_col_name = "NAME-FIRST" then worker_first_name_col = excel_col
            If formatted_col_name = "FIRSTNAME" then worker_first_name_col = excel_col
            If formatted_col_name = "NAME" then worker_full_name_col = excel_col
            If formatted_col_name = "FULLNAME" then worker_full_name_col = excel_col
            If formatted_col_name = "NAME-FULL" then worker_full_name_col = excel_col
            If formatted_col_name = "EMAIL" then worker_email_col = excel_col
            If formatted_col_name = "E-MAIL" then worker_email_col = excel_col
            If formatted_col_name = "EMAILADDRESS" then worker_email_col = excel_col
            If formatted_col_name = "E-MAILADDRESS" then worker_email_col = excel_col

            If worker_last_name_col = excel_col  Then last_name_selection  = "COLUMN " & WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) & " - " & WORKERS_COL_ARRAY(wrk_col_head, excel_col-1)
            If worker_first_name_col = excel_col Then first_name_selection = "COLUMN " & WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) & " - " & WORKERS_COL_ARRAY(wrk_col_head, excel_col-1)
            If worker_full_name_col = excel_col  Then full_name_selection  = "COLUMN " & WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) & " - " & WORKERS_COL_ARRAY(wrk_col_head, excel_col-1)
            If worker_email_col = excel_col      Then email_selection      = "COLUMN " & WORKERS_COL_ARRAY(wrk_col_letter, excel_col-1) & " - " & WORKERS_COL_ARRAY(wrk_col_head, excel_col-1)

        End If
        excel_col = excel_col + 1
    Loop until col_name = ""

    If full_name_selection <> "" Then
        first_name_selection = ""
        last_name_selection = ""
    End If

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 216, 150, "Identify Worker Information in Excel"
      DropListBox 60, 35, 150, 45, worker_columns_list, full_name_selection
      DropListBox 60, 65, 150, 45, worker_columns_list, first_name_selection
      DropListBox 60, 80, 150, 45, worker_columns_list, last_name_selection
      DropListBox 60, 110, 150, 45, worker_columns_list, email_selection
      ButtonGroup ButtonPressed
        OkButton 160, 130, 50, 15
      Text 10, 10, 205, 20, "Indicate which columns on the Excel Worksheet are the workers name and email so the script can create the lists."
      Text 20, 40, 35, 10, "Full Name:"
      Text 120, 50, 25, 10, " - OR - "
      Text 15, 70, 40, 10, "First Name:"
      Text 15, 85, 40, 10, "Last Name:"
      Text 30, 115, 25, 10, "E-Mail:"
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation

            If err_msg <> "" Then MsgBox "Please resolve to continue: " & vbNewLine & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE

    full_name_provided = TRUE
    If full_name_selection = "Select One..." Then full_name_provided = FALSE

    If full_name_provided = TRUE Then
        full_name_col_leter = left(full_name_selection, 9)
        full_name_col_leter = replace(full_name_col_leter, "COLUMN", "")
        full_name_col_leter = trim(full_name_col_leter)
    Else
        first_name_col_letter = left(first_name_selection, 9)
        first_name_col_letter = replace(first_name_col_letter, "COLUMN", "")
        first_name_col_letter = trim(first_name_col_letter)
        last_name_col_letter = left(last_name_selection, 9)
        last_name_col_letter = replace(last_name_col_letter, "COLUMN", "")
        last_name_col_letter = trim(last_name_col_letter)
    End If
    email_col_letter = left(email_selection, 9)
    email_col_letter = replace(email_col_letter, "COLUMN", "")
    email_col_letter = trim(email_col_letter)

    For each_col = 0 to UBound(WORKERS_COL_ARRAY, 2)
        this_col = WORKERS_COL_ARRAY(wrk_col_letter, each_col)
        If this_col = full_name_col_leter Then worker_full_name_col = WORKERS_COL_ARRAY(wrk_col_numb, each_col)
        If this_col = first_name_col_letter Then worker_first_name_col = WORKERS_COL_ARRAY(wrk_col_numb, each_col)
        If this_col = last_name_col_letter Then worker_last_name_col = WORKERS_COL_ARRAY(wrk_col_numb, each_col)
        If this_col = email_col_letter Then worker_email_col = WORKERS_COL_ARRAY(wrk_col_numb, each_col)

    Next

    excel_row = 2
    worker_counter = 0
    Do

        ReDim Preserve ASSIGNMENT_LISTS_ARRAY(list_message, worker_counter)
        If full_name_provided = TRUE Then
            ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, worker_counter) = trim(ObjWorkersExcel.Cells(excel_row, worker_full_name_col).Value)
        Else
            first_name = trim(ObjWorkersExcel.Cells(excel_row, worker_first_name_col).Value)
            last_name = trim(ObjWorkersExcel.Cells(excel_row, worker_last_name_col).Value)
            ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, worker_counter) = first_name & " " & last_name
        End If
        ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, worker_counter) = trim(ObjWorkersExcel.Cells(excel_row, worker_email_col).Value)
        ASSIGNMENT_LISTS_ARRAY(new_assignment, worker_counter) = TRUE

        worker_counter = worker_counter + 1
        excel_row = excel_row + 1
        next_email = trim(ObjWorkersExcel.Cells(excel_row, worker_email_col).Value)
    Loop until next_email = ""

    dlg_hgt = 75
    count_to_three = 1
    For each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY,2)
        If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) <> "" Then
            If count_to_three = 1 Then dlg_hgt = dlg_hgt + 15
            count_to_three = count_to_three + 1
            If count_to_three = 4 then count_to_three = 1
        End If
    Next

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 640, dlg_hgt, "Assignment Details"
      Text 10, 10, 440, 10, "Enter the details about the assignments to be created here:"
      Text 15, 30, 250, 10, "List of Workers to Assign Cases"
      y_pos = 45
      x_pos = 15
      for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
          If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) <> "" Then
              Text x_pos, y_pos, 200, 10, ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) & " - " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker)
              x_pos = x_pos + 205
              If x_pos = 630 Then
                  y_pos = y_pos + 15
                  x_pos = 15
              End If
          End If
      Next
      y_pos = y_pos + 20
      ButtonGroup ButtonPressed
        OkButton 530, y_pos, 50, 15
        CancelButton 585, y_pos, 50, 15
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

ElseIf worker_selection = "Manual Entry" Then
    new_lists_needed = TRUE
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 151, 60, "Number of Lists"
      EditBox 120, 10, 25, 15, number_of_assignment_lists
      ButtonGroup ButtonPressed
        OkButton 95, 35, 50, 15
      Text 10, 15, 100, 10, "How many assignment lists?"
    EndDialog

    dialog Dialog1
    cancel_without_confirmation

    ReDim Preserve ASSIGNMENT_LISTS_ARRAY(list_message, number_of_assignment_lists-1)

    dlg_hgt = 70 + number_of_assignment_lists*20
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 400, dlg_hgt, "Assignment Details"
      Text 10, 10, 440, 10, "Enter the details about the assignments to be created here:"
      Text 15, 30, 50, 10, "Worker Name"
      Text 190, 30, 50, 10, "Worker Email"
      y_pos = 45
      for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
          EditBox 15, y_pos, 150, 15, ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
          EditBox 190, y_pos, 150, 15, ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker)
          Text 345, y_pos + 5, 50, 10, "@hennepin.us"
          y_pos = y_pos + 20
          ASSIGNMENT_LISTS_ARRAY(new_assignment, each_worker) = TRUE
      next
      y_pos = y_pos + 5
      ButtonGroup ButtonPressed
        OkButton 290, y_pos, 50, 15
        CancelButton 345, y_pos, 50, 15
    EndDialog

    Do
        Do
            err_msg = ""

            dialog Dialog1
            cancel_without_confirmation

            for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
                ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) = trim(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker))
                If ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker) = "" Then err_msg = err_msg & vbNewLine & "* You must list all workers."
                If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker), " ") = 0 Then err_msg = err_msg & vbNewLine & "* Please enter a first and last name for each worker."

                ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) = trim(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker))
                If ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) = "" Then err_msg = err_msg & vbNewLine & "* Enter a worker email address for each person."
                If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker), " ") <> 0 Then err_msg = err_msg & vbNewLine & "* There are spaces in the email address, please review. For " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
                If InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker), ".") = 0 Then err_msg = err_msg & vbNewLine & "* The email address seems to be incorrect, please review. for " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
            next

            If err_msg <> "" Then MsgBox "Please resolve to continue: " & vbNewLine & vbNewLine & err_msg

        Loop until err_msg = ""
        Call check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = FALSE
End If

If new_lists_needed = TRUE Then
    If worker_selection = "Excel List" or worker_selection = "Manual Entry" or worker_selection = "Select from QI" Then
        Set objTextStream = objFSO.OpenTextFile(assignment_status_path, ForWriting, true)
    Else
        Set objTextStream = objFSO.OpenTextFile(assignment_status_path, ForAppending, true)
    End If
    for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
        If ASSIGNMENT_LISTS_ARRAY(new_assignment, each_worker) = TRUE Then
            ' ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
            space_place = InStr(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker), " ")
            first_name = left(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker), space_place-1)


            ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) = ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker) & "@hennepin.us"

            ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker) = lcase(ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker))
            ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker) = replace(ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker), " ", "-")
            ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker) = CM_mo & "-" & CM_yr & " - " & ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker) & "-" & assignment_list_title & "-assignment-list.xlsx"

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker) = "ObjExcel" & first_name

            'Write the contents of the text file
            objTextStream.WriteLine ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker) & "~" & ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker)

            'Opening the Excel file
            Set ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker) = CreateObject("Excel.Application")

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Visible = True
            Set objWorkbook = ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Workbooks.Add()
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).DisplayAlerts = True

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).ActiveSheet.Name = "Information"

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(1,1).Value = "Assigned Worker:"
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(2,1).Value = "Email:"
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(3,1).Value = "File Identifier:"
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(4,1).Value = "Total Cases:"
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(5,1).Value = "Notes:"

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Columns(1).Font.Bold = TRUE

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(1,2).Value = ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_worker)
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(2,2).Value = ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_worker)
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(3,2).Value = "ObjExcel" & first_name
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(4,2).Value = "0" 'total cases'
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(5,2).Value = "" 'notes'

            ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_worker) = 2
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Columns(1).AutoFit()
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Columns(2).AutoFit()

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Worksheets.Add().Name = "Assignment"

            assignment_column = 1
            For each_column = 0 to UBound(COLUMN_ARRAY, 2)
                If COLUMN_ARRAY(include_column, each_column) = checked Then
                    ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(1,assignment_column).NumberFormat = "@"
                    ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Cells(1,assignment_column).Value = COLUMN_ARRAY(column_header, each_column)
                    If COLUMN_ARRAY(assignment_list_col_number, each_column) = "" Then
                        COLUMN_ARRAY(assignment_list_col_number, each_column) = assignment_column
                        COLUMN_ARRAY(assignment_list_col_letter, each_column) = convert_digit_to_excel_column(assignment_column)
                        If last_assign_col < assignment_column Then last_assign_col = assignment_column
                    End If
                    assignment_column = assignment_column + 1
                End If
            Next

            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Rows(1).Font.Bold = TRUE

            for i= 1 to assignment_column
                ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).Columns(i).AutoFit()
            next

            'saving the Excel file
            ASSIGNMENT_LISTS_ARRAY(script_call_name, each_worker).ActiveWorkbook.SaveAs assignment_folder & ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_worker)
        End If
    next

    objTextStream.Close
Else
    ASSIGNMENT_LISTS_ARRAY(script_call_name, 0).worksheets("Assignment").Activate

    the_col = 1
    Do
        the_header = trim(ASSIGNMENT_LISTS_ARRAY(script_call_name, 0).Cells(1, the_col).Value)
        For each_column = 0 to UBound(COLUMN_ARRAY, 2)
            If COLUMN_ARRAY(column_header, each_column) = the_header Then
                COLUMN_ARRAY(assignment_list_col_number, each_column) = the_col
                COLUMN_ARRAY(assignment_list_col_letter, each_column) = convert_digit_to_excel_column(the_col)
                If last_assign_col < the_col Then last_assign_col = the_col
                Exit For
            End If
        Next
        the_col = the_col + 1
    Loop until the_header = ""
End If

'We fill the array here
each_line = 0
excel_row = 2
row_to_start = ""
' MsgBox "Assigned Col - " & assigned_marked_col
Do

    ReDim Preserve MASTER_LIST_ALL_ROWS(list_action, each_line)
    MASTER_LIST_ALL_ROWS(xl_col_1, each_line)   = trim(ObjExcel.Cells(excel_row, 1).Value)
    MASTER_LIST_ALL_ROWS(xl_col_2, each_line)   = trim(ObjExcel.Cells(excel_row, 2).Value)
    MASTER_LIST_ALL_ROWS(xl_col_3, each_line)   = trim(ObjExcel.Cells(excel_row, 3).Value)
    MASTER_LIST_ALL_ROWS(xl_col_4, each_line)   = trim(ObjExcel.Cells(excel_row, 4).Value)
    MASTER_LIST_ALL_ROWS(xl_col_5, each_line)   = trim(ObjExcel.Cells(excel_row, 5).Value)
    MASTER_LIST_ALL_ROWS(xl_col_6, each_line)   = trim(ObjExcel.Cells(excel_row, 6).Value)
    MASTER_LIST_ALL_ROWS(xl_col_7, each_line)   = trim(ObjExcel.Cells(excel_row, 7).Value)
    MASTER_LIST_ALL_ROWS(xl_col_8, each_line)   = trim(ObjExcel.Cells(excel_row, 8).Value)
    MASTER_LIST_ALL_ROWS(xl_col_9, each_line)   = trim(ObjExcel.Cells(excel_row, 9).Value)
    MASTER_LIST_ALL_ROWS(xl_col_10, each_line)  = trim(ObjExcel.Cells(excel_row, 10).Value)
    MASTER_LIST_ALL_ROWS(xl_col_11, each_line)  = trim(ObjExcel.Cells(excel_row, 11).Value)
    MASTER_LIST_ALL_ROWS(xl_col_12, each_line)  = trim(ObjExcel.Cells(excel_row, 12).Value)
    MASTER_LIST_ALL_ROWS(xl_col_13, each_line)  = trim(ObjExcel.Cells(excel_row, 13).Value)
    MASTER_LIST_ALL_ROWS(xl_col_14, each_line)  = trim(ObjExcel.Cells(excel_row, 14).Value)
    MASTER_LIST_ALL_ROWS(xl_col_15, each_line)  = trim(ObjExcel.Cells(excel_row, 15).Value)
    MASTER_LIST_ALL_ROWS(xl_col_16, each_line)  = trim(ObjExcel.Cells(excel_row, 16).Value)
    MASTER_LIST_ALL_ROWS(xl_col_17, each_line)  = trim(ObjExcel.Cells(excel_row, 17).Value)
    MASTER_LIST_ALL_ROWS(xl_col_18, each_line)  = trim(ObjExcel.Cells(excel_row, 18).Value)
    MASTER_LIST_ALL_ROWS(xl_col_19, each_line)  = trim(ObjExcel.Cells(excel_row, 19).Value)
    MASTER_LIST_ALL_ROWS(xl_col_20, each_line)  = trim(ObjExcel.Cells(excel_row, 20).Value)
    MASTER_LIST_ALL_ROWS(xl_col_21, each_line)  = trim(ObjExcel.Cells(excel_row, 21).Value)
    MASTER_LIST_ALL_ROWS(xl_col_22, each_line)  = trim(ObjExcel.Cells(excel_row, 22).Value)
    MASTER_LIST_ALL_ROWS(xl_col_23, each_line)  = trim(ObjExcel.Cells(excel_row, 23).Value)
    MASTER_LIST_ALL_ROWS(xl_col_24, each_line)  = trim(ObjExcel.Cells(excel_row, 24).Value)
    MASTER_LIST_ALL_ROWS(xl_col_25, each_line)  = trim(ObjExcel.Cells(excel_row, 25).Value)
    MASTER_LIST_ALL_ROWS(master_xl_row, each_line) = excel_row

    If ucase(trim(ObjExcel.Cells(excel_row, assigned_marked_col).Value)) = "TRUE" Then MASTER_LIST_ALL_ROWS(row_already_assigned, each_line) = TRUE
    If ucase(trim(ObjExcel.Cells(excel_row, assigned_marked_col).Value)) = "FALSE" Then MASTER_LIST_ALL_ROWS(row_already_assigned, each_line) = FALSE
    If trim(ObjExcel.Cells(excel_row, assigned_marked_col).Value) = "" Then MASTER_LIST_ALL_ROWS(row_already_assigned, each_line) = FALSE
    ' MsgBox "What does the col say? " & trim(ObjExcel.Cells(excel_row, assigned_marked_col).Value) & vbNewLine & "What is in the Array? " & MASTER_LIST_ALL_ROWS(row_already_assigned, each_line)

    If MASTER_LIST_ALL_ROWS(row_already_assigned, each_line) = FALSE AND row_to_start = "" Then row_to_start = excel_row

    each_line = each_line + 1
    excel_row = excel_row + 1
    next_col_one = trim(ObjExcel.Cells(excel_row, 1).Value)

Loop until next_col_one = ""

end_of_master_list = excel_row - 1
If row_to_start = "" Then row_to_start = "2"
row_to_start = row_to_start & ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 115, "Master List Row to Start"
  EditBox 95, 40, 25, 15, row_to_start
  Text 10, 10, 115, 25, "There are " & end_of_master_list & " rows in the Master List. Where would you like to start the assignments from:"
  Text 15, 45, 80, 10, "Master List row to start:"
  Text 25, 55, 95, 35, "If this is not defaulted to '2' the row to start is based on the first row on the list which 'Assigned' is not 'True'."
  ButtonGroup ButtonPressed
    OkButton 70, 95, 50, 15
EndDialog

Do
    Do
        err_msg = ""

        dialog Dialog1
        cancel_confirmation

    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

row_to_start = row_to_start * 1
assign_first_list = ""
make_even = FALSE
most_cases = 0
for each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
    which_row = ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_worker)
    If which_row > most_cases Then most_cases = which_row
next
least_cases = most_cases
For each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
    which_row = ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_worker)
    If which_row < least_cases Then least_cases = which_row
Next

For the_blank_rows = least_cases to most_cases
    For each_worker = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
        which_row = ASSIGNMENT_LISTS_ARRAY(last_excel_row, each_worker)
        If which_row =< the_blank_rows Then assign_first_list = assign_first_list & "~" & each_worker
    Next
Next

If assign_first_list <> "" Then
    assign_first_list = right(assign_first_list, len(assign_first_list)-1)
    If Instr(assign_first_list, "~") = 0 Then
        assign_first_list = array(assign_first_list)
    Else
        assign_first_list = split(assign_first_list, "~")
    End If
    make_even = TRUE
End If

last_of_assignment_lists = UBound(ASSIGNMENT_LISTS_ARRAY, 2)
If make_even = TRUE Then
    even_list_counter = 0
    last_to_make_even = UBound(assign_first_list)
End If
worker_to_assign = 0

For list_row = 0 to UBound(MASTER_LIST_ALL_ROWS, 2)
    If MASTER_LIST_ALL_ROWS(master_xl_row, list_row) >= row_to_start AND MASTER_LIST_ALL_ROWS(row_already_assigned, list_row) = FALSE Then
        If make_even = TRUE Then
            If even_list_counter > last_to_make_even Then
                worker_to_assign = 0
                make_even = FALSE
            Else
                worker_to_assign = assign_first_list(even_list_counter)
            End If
            even_list_counter = even_list_counter + 1
        End If
        excel_row = ASSIGNMENT_LISTS_ARRAY(last_excel_row, worker_to_assign)

		' MsgBox "ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign) - " & ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign)
        ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).worksheets("Assignment").Activate

        STATS_counter = STATS_counter + 1
        For master_column = 0 to UBound(COLUMN_ARRAY, 2)
            If COLUMN_ARRAY(include_column, master_column) = checked Then
                master_col_to_use = COLUMN_ARRAY(master_list_col_number, master_column)
                assign_col_to_use = COLUMN_ARRAY(assignment_list_col_number, master_column)
                ' MsgBox "Excel Row - " & excel_row & vbNewLine & "assign col - " & assign_col_to_use & vbNewLine & "master col - " & master_col_to_use & vbNewLine & "list row - " & list_row
                ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).Cells(excel_row, assign_col_to_use).Value = MASTER_LIST_ALL_ROWS(master_col_to_use, list_row)
            End If
        Next

        for i= 1 to last_assign_col
            ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).Columns(i).AutoFit()
        next

        ASSIGNMENT_LISTS_ARRAY(last_excel_row, worker_to_assign) = excel_row + 1
        ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).worksheets("Information").Activate
        ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).Cells(4, 2).Value = excel_row - 1
        ASSIGNMENT_LISTS_ARRAY(script_call_name, worker_to_assign).worksheets("Assignment").Activate

        MASTER_LIST_ALL_ROWS(row_already_assigned, list_row) = TRUE

        master_excel_row = MASTER_LIST_ALL_ROWS(master_xl_row, list_row)
        ObjExcel.Cells(master_excel_row, assigned_marked_col).Value = MASTER_LIST_ALL_ROWS(row_already_assigned, list_row)
		ObjExcel.Cells(master_excel_row, assigned_worker_col).Value = ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, worker_to_assign)

		worker_to_assign = worker_to_assign + 1
		If worker_to_assign > last_of_assignment_lists Then worker_to_assign = 0
    End If
Next
For each_assignment = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)
    ASSIGNMENT_LISTS_ARRAY(script_call_name, each_assignment).ActiveWorkbook.Save
    If close_assignment_lists_checkbox = checked Then
        ASSIGNMENT_LISTS_ARRAY(script_call_name, each_assignment).ActiveWorkbook.Close
        ASSIGNMENT_LISTS_ARRAY(script_call_name, each_assignment).Application.Quit
    End If
Next
ObjExcel.ActiveWorkbook.Save

'This part of the script will create emails to send information to the assignees if selected at the beginning of the script run.
'These emails will send automatically if an email address is known.
If send_email_checkbox = checked Then
	Call find_user_name(the_person_running_the_script)							'getting the name of the person running the script for the email signature
	For each_assignment = 0 to UBound(ASSIGNMENT_LISTS_ARRAY, 2)

		email_body = "Hello " & ASSIGNMENT_LISTS_ARRAY(assigned_worker_name, each_assignment) & ", "
		email_body = email_body & vbCr & ""
		email_body = email_body & vbCr & "Your assignment worklist has been created. It is saved in an Excel File and is ready for work now."
		email_body = email_body & vbCr & ""
		email_body = email_body & vbCr & "The assignment can be found in this Excel File: "
		email_body = email_body & vbCr & "<" & assignment_folder & ASSIGNMENT_LISTS_ARRAY(assignment_list_path, each_assignment) & ">" & vbCr
		email_body = email_body & vbCr & ""
		email_body = email_body & vbCr & "Information about this assignment: " & assignment_project_detail
		email_body = email_body & vbCr & ""
		email_body = email_body & vbCr & "If you have any questions about this assignment, please contact me."
		email_body = email_body & vbCr & ""
		email_body = email_body & vbCr & "Thank You"
		email_body = email_body & vbCr & the_person_running_the_script

		' Call create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
		send_email = True
		If ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_assignment) = "" Then send_email = False
		Call create_outlook_email(ASSIGNMENT_LISTS_ARRAY(assigned_worker_email, each_assignment), "", "Assignment List in Excel", email_body, "", send_email)

	Next
End If

script_end_procedure("Assignments complete")
