'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DISPOSABLE - BULK - Update MBI Number.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 95			 'manual run time in seconds
STATS_denomination = "M"		 'M is for each member
'END OF stats block==============================================================================================

'LOADING GLOBAL VARIABLES
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

' MsgBox "Collecting Statistics: " & collecting_statistics & vbNewLine &_
'         "Using SQL: " & using_SQL_database & vbNewLine &_
'         "Database Path: " & stats_database_path & vbNewLine &_
'         "Enhanced DB: " & STATS_enhanced_db & vbNewLine &_
'         "Collect CN: " & collect_MAXIS_case_number

'hardcode the file path for excel because this is a disposable script
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\MBI Update\MBI list.xlsx"
'Open the excel file
CALL excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)
'select the right sheet
objExcel.worksheets("Remove duplicates").Activate

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

stop_time = "4"

'Dialog to confirm the excel sheet and the worker signature
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
        BeginDialog file_select_dialog, 0, 0, 316, 110, "Select the source file"
          EditBox 5, 35, 260, 15, file_selection_path
          ButtonGroup ButtonPressed
            PushButton 270, 35, 40, 15, "Browse...", select_a_file_button
          EditBox 180, 60, 45, 15, stop_time
          ButtonGroup ButtonPressed
            OkButton 205, 85, 50, 15
            CancelButton 260, 85, 50, 15
          Text 5, 5, 305, 10, "The script has opened an Excel file to look at all the individuals that need an MBI entered. "
          Text 5, 20, 255, 10, "Check to be sure the correct file has opened and click 'Browse' if it is incorrect. "
          Text 10, 65, 160, 10, "How many hours would you like the script to run?"
          Text 10, 80, 160, 20, "Reminder, do not use Excel during the time the script is running. The script needs to use Excel."
        EndDialog

		err_msg = ""
		Dialog file_select_dialog
		If ButtonPressed = cancel then stopscript
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
            If file_selection_path = "" then
                err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
            Else
                If objExcel <> "" Then          'If there is already an excel sheet open and the browse button is pressed again - the first excel is closed and blanked out so a new one can be entered.
                    objExcel.quit
                    objExcel = ""
                End If
                call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
                err_msg = err_m & vbNewLine & "Be sure the correct Excel file opened."
            End If
		End If

		If err_msg <> "" Then MsgBox err_msg      'Display the error message
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'making stop time a number
stop_time = FormatNumber(stop_time, 2,          0,                 0,                      0)
                        'number     dec places  leading 0 - FALSE    neg nbr in () - FALSE   use deliminator(comma) - FALSE
stop_time = stop_time * 60 * 60     'tunring hours to seconds

end_time = timer + stop_time        'timer is the number of seconds from 12:00 AM so we need to add the hours to run to the time to determine at what point the script should exit the loop


'set constants for the excel columns
const case_number_col   = 1
const person_id_col     = 2
const medicare_id_col   = 3
const mbi_number_col    = 4
const first_name_col    = 5
const last_name_col     = 6
const case_status_col   = 7
const notes_col         = 8
const source_col        = 9

excel_row = 2
'Read the excel list on a loop - NO ARRAY because we are just going to act on one case at a time

'MANUAL TIME
'do - loop of all the rows in excel
end_msg = "Success! The script has completed the run and reached the end of the Excel File. The Excel File has been saved."
Do
    'MsgBox "ROW (1) is " & excel_row
    'if case number is blank - exit loop
    If trim(ObjExcel.Cells(excel_row, case_number_col).Value) = "" Then Exit Do

    'if the action indicator is done then we skip this case because it has already been acted on
    'if the action indicator is not done then we do the things
    If left(ObjExcel.Cells(excel_row, notes_col).Value, 4) <> "DONE" Then
        'set the case number from the column to this variable for nav functions to work
        'MsgBox "ROW (2) is " & excel_row
        MAXIS_case_number = ObjExcel.Cells(excel_row, case_number_col).Value
        PMI_number = ObjExcel.Cells(excel_row, person_id_col).Value
        'MsgBox "Row - " & excel_row & vbNewLine & "Case - " & MAXIS_case_number
        'MsgBox "ROW (3) is " & excel_row
        Do
            PMI_number = right(PMI_number, len(PMI_number)-1)
        Loop until left(PMI_number, 1) <> "0"
        'MsgBox "PMI: " & PMI_number

        Do
            Call navigate_to_MAXIS_screen("STAT", "SUMM")
            EMReadScreen look_for_summ, 4, 2, 46
        Loop until look_for_summ = "SUMM"
        Call back_to_SELF

        Call navigate_to_MAXIS_screen("STAT", "MEMB")

        'go through all the MEMB panels to find the right member
        Do
            EMReadScreen memb_pmi_numb, 8, 4, 46
            memb_pmi_numb = trim(memb_pmi_numb)

            If memb_pmi_numb <> PMI_number Then transmit
        Loop until memb_pmi_numb = PMI_number

        'save the member ref number for navigating
        EMReadScreen MEMB_Ref_Number, 2, 4, 33

        'MsgBox "Reference Number: " & MEMB_Ref_Number
        'go to MEDI for the selected member
        CALL navigate_to_MAXIS_screen("STAT", "MEDI")
        EMWriteScreen MEMB_Ref_Number, 20, 76
        transmit

        PF9                 'put in to edit mode
        'MsgBox "ROW (4) is " & excel_row
        'split the MBI into the three sections
        MBI_Number = trim(ObjExcel.Cells(excel_row, mbi_number_col).Value)

        MBI_one = left(MBI_Number, 4)
        MBI_two = left(right(MBI_Number, 7), 3)
        MBI_three = right(MBI_Number, 4)

        'enter each of the sections on to the MEDI panel
        EMWriteScreen MBI_one, 5, 38
        EMWriteScreen MBI_two, 5, 43
        EMWriteScreen MBI_three, 5, 47

        EMWriteScreen "M", 5, 64        'entering the source of the MBI information - which came from MMIS for this process

        transmit            'transmit to save the information to the panel
        'MsgBox "The MBI is entered"
        ' PF3                 'pf3 to send the case through background.
        '
        ' Call back_to_SELF   'go back to self
        '
        ' 'enter back in to the case
        ' CALL Navigate_to_MAXIS_screen("STAT", "MEDI")   'navigate to MEDI for the correct person.
        ' EMWriteScreen MEMB_Ref_Number, 20, 76
        ' transmit

        EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
        look_for_error = trim(look_for_error)

        If look_for_error = "WARNING:" Then     'we can transmit past warning messages and then look again
            transmit
            EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
            look_for_error = trim(look_for_error)
        End If

        If look_for_error <> "" Then        'if there is anything here - assume an error
            PF10                            'blank out the work
            If trim(ObjExcel.Cells(excel_row, notes_col).Value) = "" Then       'indicate error on the excel sheet
                ObjExcel.Cells(excel_row, notes_col).Value = "ERROR"
            ElseIf ObjExcel.Cells(excel_row, notes_col).Value = "ERROR" Then
                ObjExcel.Cells(excel_row, notes_col).Value = ObjExcel.Cells(excel_row, notes_col).Value
            Else
                ObjExcel.Cells(excel_row, notes_col).Value = "ERROR - " & ObjExcel.Cells(excel_row, notes_col).Value
            End If
            'MsgBox "ERROR"
        Else                                'if no error the number should have saved
            'Read the MBI number to be sure it succeeded.
            EMReadScreen Check_MBI_one, 4, 5, 38
            EMReadScreen Check_MBI_two, 3, 5, 43
            EMReadScreen Check_MBI_three, 4, 5, 47

            EMReadScreen Check_source, 1, 5, 64             'reading the saved source code and reformatting for the escel file
            If Check_source = "_" Then Check_source = ""

            CHECK_MBI = Check_MBI_one & Check_MBI_two & Check_MBI_three

            If CHECK_MBI = MBI_Number Then          'If it succeeded then enter 'DONE' to the action column.
                STATS_counter = STATS_counter + 1   'counting all of the people this process was completed for
                If trim(ObjExcel.Cells(excel_row, notes_col).Value) = "" Then
                    ObjExcel.Cells(excel_row, notes_col).Value = "DONE"
                ElseIf ObjExcel.Cells(excel_row, notes_col).Value = "ERROR" OR ObjExcel.Cells(excel_row, notes_col).Value = "FAILED" Then
                    ObjExcel.Cells(excel_row, notes_col).Value = "DONE"
                Else
                    ObjExcel.Cells(excel_row, notes_col).Value = "DONE - " & ObjExcel.Cells(excel_row, notes_col).Value
                End If
                'MsgBox "DONE"
            Else            'If it did not succeed then enter 'FAILED' to the action column.
                If trim(ObjExcel.Cells(excel_row, notes_col).Value) = "" Then
                    ObjExcel.Cells(excel_row, notes_col).Value = "FAILED"
                Else
                    ObjExcel.Cells(excel_row, notes_col).Value = "FAILED - " & ObjExcel.Cells(excel_row, notes_col).Value
                End If
                'MsgBox "FAILED"
            End If

            ObjExcel.Cells(excel_row, source_col).Value = Check_source          'entering the source code in to the excel file
        End If

        Call back_to_SELF   'go back to self
    ElseIf trim(ObjExcel.Cells(excel_row, source_col).Value) = "" Then          'if the excel file lists the update as DONE but the source code is not found/entered

        MAXIS_case_number = ObjExcel.Cells(excel_row, case_number_col).Value    'setting for functions
        PMI_number = ObjExcel.Cells(excel_row, person_id_col).Value
        'MsgBox "Row - " & excel_row & vbNewLine & "Case - " & MAXIS_case_number
        'MsgBox "ROW (3) is " & excel_row
        Do          'taking the leading 0s out of the pmi number
            PMI_number = right(PMI_number, len(PMI_number)-1)
        Loop until left(PMI_number, 1) <> "0"
        'MsgBox "PMI: " & PMI_number

        Do          'getting in to STAT - we do this because sometimes there are 2 people on one case and it is stuck in background
            Call navigate_to_MAXIS_screen("STAT", "SUMM")
            EMReadScreen look_for_summ, 4, 2, 46
        Loop until look_for_summ = "SUMM"
        Call back_to_SELF

        Call navigate_to_MAXIS_screen("STAT", "MEMB")       'going to find the reference number

        'go through all the MEMB panels to find the right member
        Do
            EMReadScreen memb_pmi_numb, 8, 4, 46
            memb_pmi_numb = trim(memb_pmi_numb)

            If memb_pmi_numb <> PMI_number Then transmit
        Loop until memb_pmi_numb = PMI_number

        'save the member ref number for navigating
        EMReadScreen MEMB_Ref_Number, 2, 4, 33

        'MsgBox "Reference Number: " & MEMB_Ref_Number
        'go to MEDI for the selected member
        CALL navigate_to_MAXIS_screen("STAT", "MEDI")
        EMWriteScreen MEMB_Ref_Number, 20, 76
        transmit

        EMReadScreen curr_source_code, 1, 5, 64         'reading what is on the source code for the MBI number

        If curr_source_code = "_" Then                  'if it is blank, we will update it
            'Read the MBI number to be sure it is right
            EMReadScreen Check_MBI_one, 4, 5, 38
            EMReadScreen Check_MBI_two, 3, 5, 43
            EMReadScreen Check_MBI_three, 4, 5, 47

            CHECK_MBI = Check_MBI_one & Check_MBI_two & Check_MBI_three

            MBI_Number = trim(ObjExcel.Cells(excel_row, mbi_number_col).Value)  'getting the correct number from excel

            MBI_one = left(MBI_Number, 4)
            MBI_two = left(right(MBI_Number, 7), 3)
            MBI_three = right(MBI_Number, 4)

            'MsgBox "Check MBI - " & CHECK_MBI & vbNewLine & "MBI Number - " & MBI_Number
            If CHECK_MBI = MBI_Number Then          'If the MBI is right, we will update the panel
                PF9                 'put in to edit mode

                EMWriteScreen "M", 5, 64        'entering the source of the MBI information - which came from MMIS for this process

                transmit            'transmit to save the information to the panel

                EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
                look_for_error = trim(look_for_error)

                If look_for_error = "WARNING:" Then     'we can transmit past warning messages and then look again
                    transmit
                    EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
                    look_for_error = trim(look_for_error)
                End If

                EMReadScreen Check_source, 1, 5, 64

                ObjExcel.Cells(excel_row, source_col).Value = Check_source  'entering what is actually on the panel to excel
            Else        'if the MBI is not correct, we will note the source as ? on excel for review
                ObjExcel.Cells(excel_row, source_col).Value = "?"
            End If
        Else        'if the source code is not blank, we will just enter the actual source code from the panel
            ObjExcel.Cells(excel_row, source_col).Value = curr_source_code
        End If
    End If

    'MsgBox "TIMER: " & timer
    If timer > end_time Then
        end_msg = "Success! Script has run for " & stop_time/60/60 & " hours and has finished for the time being. Excel file has been saved."
        Exit Do
    End If

    excel_row = excel_row + 1'increment to the next row.
    next_row_case = trim(ObjExcel.Cells(excel_row, case_number_col))
'loop through
Loop until next_row_case = ""

Call back_to_SELF

objWorkbook.Save

'MsgBox "Manual Time: " & STATS_counter * STATS_manualtime & vbNewLine & "For " & STATS_counter & " people." & vbNewLine & vbNewLine & "Script Time: " & timer - start_time

script_end_procedure(end_msg)
