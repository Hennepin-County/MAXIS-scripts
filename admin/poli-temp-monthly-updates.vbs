'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - POLI TEMP MONTHLY UPDATES.vbs"
start_time = timer
STATS_counter = 0                          'sets the stats counter at one
STATS_manualtime = 390                     'manual run time in seconds
STATS_denomination = "I"                   'I is for each Instance
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
call changelog_update("03/29/2024", "Updated to read from Excel list vs. entering each TE reference into a dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("02/27/2023", "Changed original procedural search month to go back 6 months vs. 2 months, and updated the file naming convention for ease of use.", "Ilse Ferris, Hennepin County")
call changelog_update("07/11/2022", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""        'Connects to BlueZone
Call check_for_MAXIS(False)
file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Knowledge Coordination\POLI TEMP\All POLI TEMP\POLI TEMP List.xlsx"
'These are two processes that will be completed. Gathering the original file, then grab the revised file.

'Initial POLI TEMP List Excel file dialog and do...loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 296, 95, "BULK - POLI TEMP TO WORD"
    GroupBox 10, 5, 280, 65, "Initial POLI TEMP File Selection"
    Text 20, 20, 260, 20, "Select the Excel file that contains the POLI TEMP references by selecting the 'Browse' button, and locating the file."
    ButtonGroup ButtonPressed
      PushButton 15, 45, 60, 15, "Browse...", select_initial_file_button
    EditBox 80, 45, 205, 15, file_selection_path
    ButtonGroup ButtonPressed
      OkButton 180, 75, 50, 15
      CancelButton 235, 75, 50, 15
EndDialog


Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_initial_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select the initial POLI TEMP List file to continue."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
Call check_for_MAXIS(False) 'Checks to make sure we're in MAXIS

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
excel_row = 2

Do
    total_code = trim(objExcel.cells(excel_row, 2).value)
    error_message = trim(objExcel.cells(excel_row, 4).value)
    If total_code = "" then exit do     'If all codes have been completed
    If error_message = "" then 
        Temp_updates = "Original,Revised"
        Temp_array = split(Temp_updates, ",")
    
        Call back_to_SELF
        For each update in temp_array
            'Setting up footer month/year based on which version we're looking at. CM - 2 since DHS will update changes in CM and CM + 1. And sometimes they report changes for one month in another month (June changes in July.)
            If update = "Original" then
                MAXIS_footer_month = CM_minus_3_mo
                MAXIS_footer_year = CM_minus_3_yr
            Elseif update = "Revised" then
                MAXIS_footer_month = CM_plus_1_mo
                MAXIS_footer_year = CM_plus_1_yr
            End if
    
            Call MAXIS_footer_month_confirmation    'confirms the footer month based on the version.
            
            Call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
            EMWriteScreen "TEMP", 5, 40     'Writes TEMP
            Call write_value_and_transmit("TABLE", 21, 71)
            policy_info = "POLI/TEMP"
        
            'Writing information and navigating in TEMP/TABLE
            Call write_value_and_transmit(total_code, 3, 21)
            EMReadScreen section_found, 18, 6, 54
            section_found = trim(section_found)
            If section_found <> total_code Then
                objExcel.cells(excel_row, 4).value = "The POLI/TEMP table reference: " & total_code & " could not be found."
            Else 
                EMReadScreen poli_title, 46, 6, 8
                poli_title = trim(poli_title)
                EmReadscreen poli_update_yr, 4, 6, 74   'This will be used to name the files
                EmReadscreen poli_update_mo, 2, 6, 79
                Call write_value_and_transmit("X", 6, 4)
            
                policy_info = policy_info & ": " & total_code
            
                'Creates the Word doc
                Set objWord = CreateObject("Word.Application")
                objWord.Visible = False 'setting visibility to false to support stabilization - poor connectivity can create an issue here.
                'sets up Word document with formatting for margins, font, title and paragraph settings.
                Set objDoc = objWord.Documents.Add()
                Set objSelection = objWord.Selection
                objSelection.PageSetup.LeftMargin = 50
                objSelection.PageSetup.RightMargin = 50
                objSelection.PageSetup.TopMargin = 30
                objSelection.PageSetup.BottomMargin = 25
                objSelection.Font.Name = "Courier New"
                objSelection.Font.Size = "14"
                objSelection.TypeText poli_title & " - "
                objSelection.TypeText policy_info
                objSelection.TypeParagraph()
                objSelection.Font.Size = "10"
                objSelection.ParagraphFormat.SpaceAfter = 0
            
                notice_length = 0
                page_nbr = 2
                'Reading TEMP reference title and information
                EMReadScreen end_of_poli, 2, 3, 79  'reading total number of reference pages
                end_of_poli = trim(end_of_poli)
                Do
                    For notice_row = 4 to 21
                        EMReadScreen poli_line, 74, notice_row, 6
                        poli_line = rtrim(poli_line)
                        If notice_row = 3 Then first_line = poli_line
                        if right(trim(poli_line),9) = "FMINFO___" Then poli_line = ""
                        If right(trim(poli_line),4) = "Page" Then
                            poli_line = rtrim(poli_line) & " " & page_nbr
                            page_nbr = page_nbr + 1
                        End If
                        poli_wording = poli_wording & poli_line & vbcr
                        If left(trim(poli_line), 7) = "WORKER:" Then Exit For
                        poli_line = ""
                    Next
                    EMReadScreen current_page, 2, 3, 72
                    current_page = trim(current_page)
                    PF8
                    notice_length = notice_length + 1
                Loop until current_page = end_of_poli
        
                objSelection.TypeText poli_wording  'exporting temp verbiage to Word
                'adding closing message to document about when information was collected.
                objSelection.TypeParagraph()
            
                '----------------------------------------------------------------------------------------------------File information coding
                If right(poli_title, 1) = "." then poli_title = left(poli_title, len(poli_title) - 1) 'sometimes there is an extra period in the title.
            
                'These characters will not allow the file to save. Replacing them based on the character found.
                poli_title = replace(poli_title, ":", " ")
                poli_title = replace(poli_title, "/", " ")
                poli_title = replace(poli_title, "?", " ")
                poli_title = replace(poli_title, "<", "Under ")
                poli_title = replace(poli_title, chr(34), "")   'chr(34) is ""
            
                'folder paths
                compare_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Knowledge Coordination\POLI TEMP\Comparison Files"
                diff_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Knowledge Coordination\POLI TEMP\" 'KC folder where the DIFF files will be housed
            
                'Creating file names
                poli_update_date = " " & poli_update_yr & " - " & poli_update_mo
                file_name = "\" & total_code & " " & new_poli & poli_title & poli_update_date & ".docx"
                If update = "Original" then original_file = file_name   'changes the file name to be compared below
                If update = "Revised" then revised_file = file_name     'ditto
            
                objDoc.SaveAs(compare_file_path & file_name)
                objWord.Visible = True  'Setting visibility back to true prior to quit. Does not need to be before the save.
                objWord.Quit
            
                'blanking out the variables
                'total_code = ""
                poli_title = ""
                poli_wording = ""
                policy_info = ""
                poli_update_date = ""
                file_name = ""
            End if
        Next  
        '----------------------------------------------------------------------------------------------------Comparing the two files and creating a new file to be saved w/ changes tracked.
        If trim(objExcel.cells(excel_row, 4).value) = "" then
            'Creating single variable to compare below
            old_poli_file = compare_file_path & original_file
            new_poli_file = compare_file_path & revised_file
                    
            Set objWord = CreateObject("Word.Application")  'set application object
            objWord.Documents.Open old_poli_file            'opening old file - original temp file
            objWord.ActiveDocument.Compare new_poli_file    'comparing the new file - revised temp file
            objWord.Visible = True
                    
            Set objDoc = objWord.ActiveDocument             'set document object
            objDoc.SaveAs(diff_file_path & revised_file)
            objWord.Quit
                
            STATS_counter = STATS_counter + 1
            Call File_Exists(new_poli_file, does_file_exist)
            If does_file_exist = True then objExcel.cells(excel_row, 4).value = "Diff file created."
        End if 
    End if 
    excel_row = excel_row + 1
Loop until trim(objExcel.cells(excel_row, 2).value) = ""

script_end_procedure("Success!!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------07/14/2022
'--Tab orders reviewed & confirmed----------------------------------------------07/14/2022
'--Mandatory fields all present & Reviewed--------------------------------------07/14/2022
'--All variables in dialog match mandatory fields-------------------------------07/14/2022
'Review dialog names for content and content fit in dialog----------------------02/27/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)----------------------------------07/14/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky-------------------------------------------07/14/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable------------------------------------07/14/2022-------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-07/14/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------07/14/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------07/14/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------07/14/2022-------------------N/A
'--Out-of-County handling reviewed----------------------------------------------07/14/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------07/14/2022
'--BULK - review output of statistics and run time/count (if applicable)--------07/14/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---07/14/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------07/14/2022
'--Incrementors reviewed (if necessary)-----------------------------------------12/03/2023
'--Denomination reviewed -------------------------------------------------------07/14/2022
'--Script name reviewed---------------------------------------------------------07/14/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------12/03/2023-------------------N/A
'
'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------07/14/2022
'--comment Code-----------------------------------------------------------------07/14/2022
'--Update Changelog for release/update------------------------------------------02/27/2023
'--Remove testing message boxes-------------------------------------------------07/14/2022
'--Remove testing code/unnecessary code-----------------------------------------07/14/2022
'--Review/update SharePoint instructions----------------------------------------07/14/2022-------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------07/14/2022-------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------02/27/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------02/27/2023-------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------07/14/2022-------------------N/A
'--Update project team/issue contact (if applicable)----------------------------07/14/2022
