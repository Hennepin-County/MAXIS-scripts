'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - BULK POLI TEMP TO WORD.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
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
call changelog_update("09/29/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function select_folder(folder_selected)
''--- This function opens a "Select Folder" dialog and will return the fully qualified path of the selected folder.
''~~~~~ folder_selected: variable for the name of the file
''===== Keywords: MAXIS, MMIS, PRISM, folder
    Dim objFolder, objItem, objShell
    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, folder_selected)
    If IsObject(objfolder) Then folder_selected = objFolder.Self.Path     ' Return the path of the selected folder
End Function

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""        'Connects to BlueZone
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

get_county_code

If worker_county_code = UCase("X127") then
    file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\Knowledge Coordination\POLI TEMP\All POLI TEMP\POLI TEMP List.xlsx"
    root_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\Knowledge Coordination\POLI TEMP\All POLI TEMP\" 'KC folder where the DIFF files will be housed
    Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
Else
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
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_initial_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select the initial POLI TEMP List file to continue."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 296, 100, "BULK - POLI TEMP TO WORD"
        GroupBox 10, 5, 280, 70, "File Where the POLI TEMP Word Docs Will Be Saved"
        Text 15, 20, 275, 20, "Select the folder where the script will save the Word Documents by selecting the 'Browse' button, and locating the folder."
        ButtonGroup ButtonPresse
          PushButton 15, 45, 60, 15, "Browse...", select_save_folder_button
        EditBox 80, 45, 205, 15, root_file_path
        ButtonGroup ButtonPressed
          OkButton 180, 80, 50, 15
          CancelButton 235, 80, 50, 15
    EndDialog

    Do 
        Do
            err_msg = ""
            dialog Dialog1
            cancel_without_confirmation
            If ButtonPressed = select_save_folder_button then call select_folder(root_file_path)
            'adds in formatting for the root file for later saving purposes 
            If trim(root_file_path) = "" then
                err_msg = err_msg & vbcr & "* Select the folder where to Word documents will be saved to continue."
            Else
                If right(root_file_path, 1) <> "\" or right(root_file_path, 1) <> "/" then root_file_path = root_file_path & "\"
            End if
            IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
        Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
End If

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Adding Excel column for file creation confirmation
objExcel.Cells(1, 4).Value = "POLI TEMP File Found in Folder"

FOR i = 1 to 4		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Call check_for_MAXIS(False) 'Checks to make sure we're in MAXIS
Call MAXIS_footer_month_confirmation    'confirms the footer month based on the version.
excel_row = 2

Do
    Call back_to_SELF
    total_code = trim(objExcel.cells(excel_row, 2).value)
    Call navigate_to_MAXIS_screen("POLI", "____")   'Navigates to POLI (can't direct navigate to TEMP)
    EMWriteScreen "TEMP", 5, 40     'Writes TEMP
    Call write_value_and_transmit("TABLE", 21, 71)

    policy_info = "POLI/TEMP"
    panel_title = "TABLE"
    'Writing information and navigating in TEMP/TABLE
    Call write_value_and_transmit(total_code, 3, 21)
    EMReadScreen section_found, 18, 6, 54
    section_found = trim(section_found)
    If section_found = total_code Then
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
        objSelection.Font.Name = "Montserrat"
        objSelection.Font.Size = "14"
        objSelection.TypeText poli_title & " - "
        objSelection.TypeText policy_info
        objSelection.TypeParagraph()
        objSelection.Font.Size = "12"
        objSelection.ParagraphFormat.SpaceAfter = 0

        notice_length = 0
        page_nbr = 2
        'Reading TEMP reference title and information
        EMReadScreen end_of_poli, 2, 3, 79  'reading total number of reference pages
        end_of_poli = trim(end_of_poli)
        Do
            For notice_row = 4 to 21
                EMReadScreen poli_line, 74, notice_row, 6
                poli_line = trim(poli_line)
                If notice_row = 3 Then first_line = poli_line
                if right(trim(poli_line),9) = "FMINFO___" Then poli_line = ""
                If right(trim(poli_line),4) = "Page" Then
                    poli_line = trim(poli_line) & " " & page_nbr
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

        '----------------------------------------------------------------------------------------------------File information coding
        If right(poli_title, 1) = "." then poli_title = left(poli_title, len(poli_title) - 1) 'sometimes there is an extra period in the title.
        'These characters will not allow the file to save. Replacing them based on the character found.
        poli_title = replace(poli_title, ":", " ")
        poli_title = replace(poli_title, "/", " ")
        poli_title = replace(poli_title, "?", " ")
        poli_title = replace(poli_title, "<", "Under ")
        poli_title = replace(poli_title, chr(34), "")   'chr(34) is ""

        temp_code = "TE " & right(total_code, len(total_code) -2) 'saving convention to save POLI TEMP code as TE xx.xx.xx vs. TExx.xx.xx (creating space to make more searchable in SPO)

        'folder paths and saving each document
        poli_file_name = root_file_path & temp_code & " " & poli_title & ".pdf"
        'The number '17' is a Word Ennumeration that defines this should be saved as a PDF.
        objWord.Visible = True  'Setting visibility back to true prior to quit. Ooes not need to be before the save.
        objDoc.SaveAs poli_file_name, 17
        objDoc.Close wdDoNotSaveChanges     'This allows us to close without any changes to the Word Document. Since we have the PDF we do not need the Word Doc
        objWord.Quit						'close Word Application instance we opened. (any other word instances will remain)
        
        'blanking out the variables
        total_code = ""
        poli_title = ""
        poli_wording = ""
        policy_info = ""
        poli_update_date = ""
        file_name = ""
    End if

    Call File_Exists(poli_file_name, does_file_exist)   'This bit just makes sure that the file actually exists in the folder.
    objExcel.cells(excel_row, 4).value = does_file_exist
    excel_row = excel_row + 1
    STATS_counter = STATS_counter + 1
Loop until trim(objExcel.cells(excel_row, 2).value) = ""

STATS_counter= STATS_counter - 1
script_end_procedure("Success!!")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------07/24/2023
'--Tab orders reviewed & confirmed----------------------------------------------07/24/2023
'--Mandatory fields all present & Reviewed--------------------------------------07/24/2023
'--All variables in dialog match mandatory fields-------------------------------07/24/2023
'Review dialog names for content and content fit in dialog----------------------07/24/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------07/24/2023-----------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------07/24/2023-----------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------07/24/2023-----------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used--07/24/2023-----------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------07/24/2023
'--MAXIS_background_check reviewed (if applicable)-------------------------------07/24/2023-----------------N/A
'--PRIV Case handling reviewed --------------------------------------------------07/24/2023-----------------N/A
'--Out-of-County handling reviewed-----------------------------------------------07/24/2023-----------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------07/24/2023
'--BULK - review output of statistics and run time/count (if applicable)--------07/24/2023
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------07/24/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------07/24/2023
'--Incrementors reviewed (if necessary)-----------------------------------------07/24/2023
'--Denomination reviewed -------------------------------------------------------07/24/2023
'--Script name reviewed---------------------------------------------------------07/24/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------07/24/2023

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------07/24/2023
'--comment Code-----------------------------------------------------------------07/24/2023
'--Update Changelog for release/update------------------------------------------07/24/2023
'--Remove testing message boxes-------------------------------------------------07/24/2023
'--Remove testing code/unnecessary code-----------------------------------------07/24/2023
'--Review/update SharePoint instructions----------------------------------------Instructions incoming
'--Other SharePoint sites review (HSR Manual, etc.)------------------------------07/24/2023-----------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------07/24/2023
'--COMPLETE LIST OF SCRIPTS update policy references-----------------------------07/24/2023-----------------N/A
'--Complete misc. documentation (if applicable)---------------------------------07/24/2023
'--Update project team/issue contact (if applicable)----------------------------07/24/2023
