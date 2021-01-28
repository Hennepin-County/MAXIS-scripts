'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - GET MEMBER SSN.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
call changelog_update("01/03/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\VA COLA SSN.xlsx"

'dialog and dialog DO...Loop	
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed 
		Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 221, 50, "Select the UNEA income source file"
  			ButtonGroup ButtonPressed
    		PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    		OkButton 110, 30, 50, 15
    		CancelButton 165, 30, 50, 15
  			EditBox 5, 10, 165, 15, file_selection_path
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

'Sets up the array to store all the information for each client'
Dim COLA_array()
ReDim COLA_array (7, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num    	= 1			'Each of the case numbers will be stored at this position'
Const first_name	= 2
Const clt_SSN    	= 3
Const memb_num		= 4
Const inc_type		= 5
Const claim_num   	= 6
Const unea_amt 	  	= 7

'Now the script adds all the clients on the excel list into an array
Excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
	member_first_name = objExcel.cells(excel_row, 4).value	'establishes client SSN
	
	MAXIS_case_number = trim(MAXIS_case_number)
	If MAXIS_case_number = "" then exit do
	
	'Adding client information to the array'
	ReDim Preserve COLA_array(7, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	COLA_array (case_num, 	entry_record) = MAXIS_case_number		'The client information is added to the array'
	COLA_array (first_name, entry_record) = trim(member_first_name)
	COLA_array (clt_SSN,  	entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	Stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop

back_to_self

For i = 0 to Ubound(COLA_array, 2)
	'Establishing values for each case in the array of cases 
	MAXIS_case_number = COLA_array (case_num, i)	
	member_first_name = COLA_array (first_name, i)
	If trim(MAXIS_case_number) = "" then exit for
 	
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		COLA_array (clt_SSN,  	entry_record) = ""
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
	Else 
		row = 5
		HH_count = 0
		Do 
			EMReadScreen client_name, 12, 6, 63
			client_name = replace(client_name, "_", "")
			'msgbox client_name & COLA_array(first_name, i)
			If client_name = COLA_array (first_name, i) then 
				EMReadScreen full_SSN, 11, 7, 42
				full_SSN = replace(full_SSN, " ", "-")
				'msgbox full_SSN
				COLA_array (clt_SSN,i) = full_SSN 
				matched_SSN = True 
				EMReadScreen member_number, 2, 4, 33
				COLA_array (memb_num,i) = member_number
				exit do
			ELSE
			 	matched_SSN = False
				HH_count = HH_count + 1
				transmit
				EMReadScreen MEMB_error, 5, 24, 2
			End if 
		Loop until MEMB_error = "ENTER"
		
		IF matched_SSN = False then 
			COLA_array (clt_SSN, i) = ""
		Else 
			Call navigate_to_MAXIS_screen("STAT", "UNEA")
			Call write_value_and_transmit(member_number, 20, 76)
					
			EMReadScreen total_amt_of_panels, 1, 2, 78	'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
			If total_amt_of_panels = "0" then 
				COLA_array (inc_type, i) = ""
				COLA_array (claim_num, i) = ""
				COLA_array (unea_amt, i) = ""
			Else 	
				Do
					EMReadScreen current_panel_number, 1, 2, 73
					EMReadScreen income_type, 2, 5, 37
					If income_type = "11" or income_type = "12" or income_type = "13" or income_type = "38" then
						income_panel_found = true 
						EMReadScreen claim_number, 15, 6, 37
						EMReadScreen unea_amount, 8, 13, 68
						
						COLA_array (inc_type, i) = income_type
						COLA_array (claim_num, i) = replace(claim_number, "_", "")
						COLA_array (unea_amt, i) = replace(unea_amount, "_", "")
						exit do
					Else 
						transmit
					End if 
				Loop until current_panel_number = total_amt_of_panels
			End if 
		End if 
	END if 
Next 
		
'Export data to Excel 
Excel_row = 2
For i = 0 to Ubound(COLA_array, 2)
	ObjExcel.Cells(Excel_row, 5).Value = COLA_array(clt_SSN, i)
	ObjExcel.Cells(Excel_row, 6).Value = COLA_array(memb_num, i)
	ObjExcel.Cells(Excel_row, 10).Value = COLA_array(inc_type, i)
	ObjExcel.Cells(Excel_row, 11).Value = COLA_array(claim_num, i)
	ObjExcel.Cells(Excel_row, 13).Value = COLA_array(unea_amt, i)
	Excel_row = Excel_row + 1
Next

FOR i = 1 to 15		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Stats_counter = stats_counter - 1
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")