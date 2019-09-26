'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - AVS FORMS PROCESSING.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "C"       			   'M is for each CASE
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
call changelog_update("09/24/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog , 0, 0, 266, 115, "AVS Forms Procesing"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used to determine if AVS forms have been rec'd for a recipient in ECF."
  Text 15, 70, 230, 15, "Select the Excel file that contains the ECF info by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\AVS\New Mail Workflow 09-23-2019 ECF.xlsx"

Do
    err_msg = ""
	dialog
	cancel_without_confirmation 
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""
If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

excel_row = 2
entry_record = 0

DIM master_array()
ReDim master_array(2, 0)

const SMI_ECF_const   = 0
const scan_date_const = 1

Do 
	SMI_ECF_number  = ObjExcel.Cells(excel_row, 1).Value
	SMI_ECF_number  = trim(SMI_ECF_number)
    If SMI_ECF_number = "" then exit do 
    
    scan_date = ObjExcel.Cells(excel_row, 3).Value
    scan_date = trim(scan_date)
    
    ReDim Preserve master_array(2, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    master_array(smi_number_const,	entry_record) = SMI_ECF_number 		
    master_array(scan_date_const, 	entry_record) = scan_date 				
    
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
LOOP

msgbox entry_record

objExcel.Quit   'Closes the initial spreadsheet 
objExcel = ""

'----------------------------------------------------------------------------------------------------GATHERING THE LIST OF ALL BANKED MONTHS CASES 
'dialog and dialog DO...Loop	
BeginDialog , 0, 0, 266, 115, "AVS Master List Selection"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection
  Text 20, 20, 235, 25, "Now select the master AVS list to start the filter process."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\AVS\AVS Forms Distribution Master List.xlsx"

Do
    err_msg = ""
	dialog
	cancel_without_confirmation 
	If ButtonPressed = select_file_button then call file_selection_system_dialog(file_selection, ".xlsx")
    If trim(file_selection) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""
If objExcel = "" Then call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
'----------------------------------------------------------------------------------------------------FILTERING THE ARRAY 
form_count = 0
excel_row = 2

DO 
    SMI_number = ObjExcel.Cells(excel_row, 7).Value
    SMI_number = trim(SMI_number)
    If SMI_number = "" then exit do 
    
    For item = 0 to UBound(master_array, 2)
        SMI_ECF_number = master_array(SMI_ECF_const, item)  
        scan_date = master_array(scan_date_const, item)
        
        If SMI_ECF_number = SMI_number then 
            match_found = True 
            objExcel.Cells(excel_row, 18).Value = scan_date
            form_count = form_count + 1
            'msgbox scan_date
            exit for
        else 
            match_found = False 
        end if 
    Next

    excel_row = excel_row + 1
    
Loop 

msgbox form_count    
STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your AVS master list.")