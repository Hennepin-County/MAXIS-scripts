'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - EXPEDITED REVIEW NOTES GATHERING.vbs"
start_time = timer
STATS_counter = 1                           'sets the stats counter at one
STATS_manualtime = 30                       'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
CALL changelog_update("07/01/2020", "Phase two notes gathering from previous days' assignment updated.", "Ilse Ferris, Hennepin County")
CALL changelog_update("06/19/2020", "Updated code in prep for next phase", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/09/2020", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.

changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 261, 65, "ADMIN - EXPEDITED REVIEW NOTES GATHERING"
  ButtonGroup ButtonPressed
    OkButton 150, 45, 50, 15
    CancelButton 205, 45, 50, 15
  Text 20, 20, 220, 20, "This script should be used to pull in assignment notes from QI team before running the current day's report for EXP SNAP."
  GroupBox 10, 5, 250, 35, "Using this script:"
EndDialog

dialog Dialog1
cancel_without_confirmation 

'Establshing array     
DIM master_array()          'Declaring the array
ReDim master_array(2, 0)    'Resizing the array 
    
'Creating constants to value the array elements
const master_case_number_const = 0  
const master_notes_const       = 1
    
master_record = 0    'incrementer for the array 


array_of_assigments = array("Expedited Review ","Pending Over 30 Days ")

For each assignment in array_of_assigments
    assignment = replace(assignment, """","")
    
    previous_date = dateadd("d", -1, date)
    Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the file names
    file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
    file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & assignment & file_date & ".xlsx"
    
    If objExcel = "" Then Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
    excel_row = 2
    Do 
        master_case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
        master_case_number = trim(master_case_number)
        If master_case_number = "" then exit do 
        
        master_notes = ObjExcel.Cells(excel_row, 8).Value       'reading worker entered notes       
    
        If trim(master_notes) <> "" then 
            ReDim Preserve master_array(2, master_record)	'This resizes the array based on if master notes were found or not
            master_array(master_case_number_const,  master_record) = master_case_number
            master_array(master_notes_const,        master_record) = trim(master_notes)
            
            master_record = master_record + 1			'This increments to the next entry in the array'
            STATS_counter = STATS_counter + 1           'stats incrementor 
        End if 
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP
        
    'Closing workbook and quiting Excel application
    objExcel.ActiveWorkbook.Close                           
    objExcel.Application.Quit
    objExcel.Quit
    objExcel = ""
Next     

'Finding and opening the previous day's file. 
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the fin
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
previous_file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & file_date & ".xlsx"

If objExcel = "" Then Call excel_open(previous_file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "Report 1" then sheet_list = sheet_list & objWorkSheet.Name & ","
Next
    
sheet_list = trim(sheet_list)  'trims excess spaces of sheet_list
If right(sheet_list, 1) = "," THEN sheet_list = left(sheet_list, len(sheet_list) - 1) 'trimming off last comma
array_of_sheets = split(sheet_list, ",")   'Creating new array
    
For each excel_sheet in array_of_sheets
    objExcel.worksheets(excel_sheet).Activate 'Activates the applicable worksheet 
    excel_row = 2
    
    Do 
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value  'reading case number
        MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do 
            
        For i = 0 to Ubound(master_array, 2)                                                            'If notes were selected to be added, array is looped thru for matching case number
            master_case_number = master_array(master_case_number_const, i)
            If master_case_number = MAXIS_case_number then 
                ObjExcel.Cells(excel_row, 8).Value = master_array(master_notes_const, i)   'If case number is found, prevoius list notes are added to the array 
                exit for 
            End if      
        Next 
          
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP
Next 
    
'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the script run is complete! Count of notes = " & master_record)