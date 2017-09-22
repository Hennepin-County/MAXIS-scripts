'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - UPDATE BM CASE REVIEW LIST.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'default file path since this is stationary
file_selection_path = "\\hcgg.fr.co.hennepin.mn.us\lobroot\HSPH\Team\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Master case review list.xlsx"

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			BeginDialog file_select_dialog, 0, 0, 226, 50, "Select the banked months case review file."
  				ButtonGroup ButtonPressed
    			PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    			OkButton 110, 30, 50, 15
    			CancelButton 165, 30, 50, 15
  				EditBox 5, 10, 165, 15, file_selection_path
			EndDialog
			err_msg = ""
			Dialog file_select_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
			End If
			If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
	
BeginDialog update_banked_month_status_dialog, 0, 0, 191, 60, "Dialog"
  DropListBox 80, 10, 105, 15, "Select one..."+chr(9)+"January"+chr(9)+"February"+chr(9)+"March"+chr(9)+"April"+chr(9)+"May"+chr(9)+"June"+chr(9)+"July"+chr(9)+"August"+chr(9)+"September"+chr(9)+"October"+chr(9)+"November"+chr(9)+"December", month_selection
  ButtonGroup ButtonPressed
    OkButton 80, 35, 50, 15
    CancelButton 135, 35, 50, 15
  Text 5, 15, 70, 10, "Update status month:"
EndDialog

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog update_banked_month_status_dialog
		If ButtonPressed = 0 then StopScript
		If month_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select the status month to update."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in			
	
'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

'starts adding phone numbers at row selected
Select Case month_selection
Case "January"
	MAXIS_footer_month = "01"
	MAXIS_footer_year = "17"
	excel_col = 17
Case "February"
	MAXIS_footer_month = "02"
	MAXIS_footer_year = "17"
	excel_col = 18
Case "March"
	MAXIS_footer_month = "03"
	MAXIS_footer_year = "17"
	excel_col = 19
Case "April"
	MAXIS_footer_month = "04"
	MAXIS_footer_year = "17"
	excel_col = 20
Case "May"
	MAXIS_footer_month = "05"
	MAXIS_footer_year = "17"
	excel_col = 21
Case "June"
	MAXIS_footer_month = "06"
	MAXIS_footer_year = "17"
	excel_col = 22
Case "July"
	MAXIS_footer_month = "07"
	MAXIS_footer_year = "17"
	excel_col = 23
Case "August"
	MAXIS_footer_month = "08"
	MAXIS_footer_year = "17"
	excel_col = 24
Case "September"
	MAXIS_footer_month = "09"
	MAXIS_footer_year = "17"
	excel_col = 25
Case "October"
	MAXIS_footer_month = "10"
	MAXIS_footer_year = "17"
	excel_col = 26
Case "November"
	MAXIS_footer_month = "11"
	MAXIS_footer_year = "17"
	excel_col = 27
Case "December"
	MAXIS_footer_month = "12"
	MAXIS_footer_year = "17"
	excel_col = 28
End Select

excel_row = 2
DO  
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 3).value
    If MAXIS_case_number = "" then exit do
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
	
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then msgbox MAXIS_case_number & " cannot access CASE/CURR."
    
    EMReadScreen case_status, 8, 8, 9
    case_status = trim(case_status)
    If case_status = "INACTIVE" then 
        ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
	Elseif case_status = "ACTIVE" then 
        MAXIS_row = 9
        Do 
            EMReadScreen prog_name, 4, MAXIS_row, 3
            prog_name = trim(prog_name)
            if prog_name = "" then exit do
            If prog_name = "FS" then 
                EMReadScreen case_status, 8, MAXIS_row, 9
                case_status = trim(case_status)
	            if case_status = "ACTIVE" then 
                    exit do
                ELSE 
                    MAXIS_row = MAXIS_row + 1
                END IF 
            Else
                MAXIS_row = MAXIS_row + 1
            END IF 
	    Loop until MAXIS_row = 17
        If prog_name <> "FS" then ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
    END If 
	
	'inputs EXEMPT on cases that are active on GA/open on other programs
	If case_status = "ACTIVE" then 
		Call navigate_to_MAXIS_screen("STAT", "PROG")
		EMReadScreen SNAP_prog_status, 4, 10, 74 
		IF SNAP_prog_status = "INAC" or trim(SNAP_prog_status) = "" or SNAP_prog_status = "DENY" then 
			ObjExcel.Cells(excel_row, excel_col).Value = "Inactive"
		Elseif SNAP_prog_status = "PEND" then 
			ObjExcel.Cells(excel_row, excel_col).Value = "Pending"
		Else 
			EMReadScreen cash_prog_status, 4, 6, 74
			IF cash_prog_status = "ACTV" then 
				ObjExcel.Cells(excel_row, excel_col).Value = "Exempt"
			Else
				ObjExcel.Cells(excel_row, excel_col).Value = ""
			End if 
		END IF  
	END IF

    MAXIS_case_number = ""
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 3).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The Excel file now has been update for all inactive SNAP cases.")
