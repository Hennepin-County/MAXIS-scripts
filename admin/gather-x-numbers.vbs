''STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - GATHER X NUMBERS.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "20"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("07/10/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\X numbers.xlsx"

'dialog and dialog DO...Loop
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed
            Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 266, 110, "X number dialog"
  				ButtonGroup ButtonPressed
    			PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    			OkButton 145, 90, 50, 15
    			CancelButton 200, 90, 50, 15
  				EditBox 15, 45, 180, 15, file_selection_path
  				GroupBox 10, 5, 250, 80, "Using the GATHER X NUMBERS script"
  				Text 20, 20, 235, 20, "This script should be used when updating worker information to be used later in scripts or otherwise."
  				Text 15, 65, 230, 15, "Select the Excel file that contains the X number information by selecting the 'Browse' button, and finding the file."
			EndDialog

			err_msg = ""

			Dialog Dialog1
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

'Gathering case status for answered call cases
objExcel.worksheets("Worker X numbers").Activate

call navigate_to_MAXIS_screen("rept", "user")		'Getting to REPT/USER
PF5													'Hitting PF5 to force sorting, which allows directly selecting a county
EMWriteScreen county_code, 21, 6					'Inserting county
transmit

excel_row = 2
row = 7												'Declaring the MAXIS row
Do
	Do
		'Reading MAXIS information for this row, adding to spreadsheet
		EMReadScreen worker_ID, 8, row, 5			'worker ID
		EMReadScreen worker_name, 14, row, 14
		If trim(worker_ID) = "" then exit do		'exiting before writing to array, in the event this is a blank (end of list)
		EMReadScreen phone_number, 12, row, 69
		If trim(phone_number) = "" then exit do

		If instr(worker_name, "HENN CO") then
		 	add_to_excel = False
		ElseIf instr(worker_name, "HENNEPIN COUNTY") then
		 	add_to_excel = False
		elseIf instr(worker_name, "HSPH") then
		 	add_to_excel = False
		elseIf instr(worker_name, "INACTIVE") then
		 	add_to_excel = False
		elseIf instr(worker_name, "INACTV") then
		 	add_to_excel = False
		elseIf instr(worker_name, "MAXIS") then
		 	add_to_excel = False
		ElseIf instr(worker_name, "TESTER") then
		 	add_to_excel = False
		elseIf instr(worker_name, "TESTING") then
		 	add_to_excel = False
		else
			add_to_excel = true
			ObjExcel.Cells(excel_row, 1).Value = worker_ID
			ObjExcel.Cells(excel_row, 2).Value = worker_name
			excel_row = excel_row + 1
			STATS_counter = STATS_counter + 1
		End if
		worker_ID = ""
		worker_name = ""
		row = row + 1
	Loop until row = 19

	'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
	EMReadScreen more_pages_check, 7, 19, 3
	If more_pages_check = "More: +" then
		PF8			'getting to next screen
		row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
	End if
Loop until trim(more_pages_check) = "More:" or trim(more_pages_check) = ""	'The or works because for one-page only counties, this will be blank

FOR i = 1 to 2		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1
'msgbox STATS_counter
script_end_procedure("Success, your list is complete!")
