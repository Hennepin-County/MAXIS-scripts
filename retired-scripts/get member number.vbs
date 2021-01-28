'Required for statistical purposes===============================================================================
name_of_script = "BULK - SEND CBO MANUAL REFERRALS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 120                     'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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
call changelog_update("02/13/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone and establishing county name
EMConnect ""	

'dialog and dialog DO...Loop	
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed 
			Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 266, 110, "GRH case case numbers"
  				ButtonGroup ButtonPressed
    			PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 90, 50, 15
    			CancelButton 200, 90, 50, 15
  				EditBox 15, 45, 180, 15, file_selection_path
  				GroupBox 10, 5, 250, 80, ""
  				Text 20, 20, 235, 20, ""
  				Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
			EndDialog
			err_msg = ""
			Dialog Dialog1
            cancel_without_confirmation
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

'Ensures that user is in current month
back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

'Now the script adds all the clients on the excel list into an array for the appropriate county
excel_row = 2 're-establishing the row to start checking the members for

'Gathering info from MAXIS, and making the referrals and case notes if cases are found and active----------------------------------------------------------------------------------------------------
Do 
    MAXIS_case_number = objExcel.cells(excel_row, 1).Value
    
    MAXIS_case_number = Trim(MAXIS_case_number)
    If trim(MAXIS_case_number) = "" then exit do 
    
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
    row = 5
    HH_count = 0
    Do 
        EMReadScreen member_number, 2, row, 3
        HH_count = HH_count + 1
        transmit
        EMReadScreen MEMB_error, 5, 24, 2
    Loop until MEMB_error = "ENTER"
    If HH_count = 1 then 
        ObjExcel.Cells(excel_row, 2).Value = member_number 
    Else 
        ObjExcel.Cells(excel_row, 2).Value = "add manually" 
    End if 

    STATS_counter = stats_counter + 1
    excel_row = excel_row + 1    
    maxis_case_number = ""
Loop 	

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! Review the spreadsheet for accuracy. Some cases may not have had a referral made.")