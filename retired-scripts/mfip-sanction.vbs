	'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MFIP SANCTION.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "150"                'manual run time in seconds
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
call changelog_update("06/12/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 125, "MFIP sanction"
  EditBox 75, 20, 135, 15, worker_number
  CheckBox 5, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 105, 105, 50, 15
    CancelButton 160, 105, 50, 15
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 40, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 5, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
EndDialog
'Shows dialog
Do
	Do
		Dialog Dialog1 
		Cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(2,  1).Value = "Worker"
ObjExcel.Cells(2,  2).Value = "Case Number"
ObjExcel.Cells(2,  3).Value = "Client Name"
ObjExcel.Cells(2,  4).Value = "Memb #"
ObjExcel.Cells(2,  5).Value = "Vendor Rsn"
ObjExcel.Cells(2,  6).Value = "ABPS?"       'Y or N
ObjExcel.Cells(2,  7).Value = "Attended FO?"
objExcel.cells(2,  8).Value = "Orient Date"
ObjExcel.Cells(2,  9).Value = "EMPS Sanc Rsn"
objExcel.cells(2, 10).Value = "Begin Date"
objExcel.cells(2, 11).Value = "End Date"
objExcel.cells(2, 12).Value = "Last 2 months"   'TIME CODE in the last 2 months
objExcel.cells(2, 13).Value = "60 mo. EXT RSN"
objExcel.cells(2, 14).Value = "Total TANF mo."
objExcel.cells(2, 15).Value = "Total # Sanctions"
ObjExcel.Cells(2, 16).Value = "Sanc %"
objExcel.cells(2, 17).Value = "SANC Last 6 months"   'Any SANC codes in the last 6 months
objExcel.cells(2, 18).Value = "Date closed 7th"
objExcel.cells(2, 19).Value = "Date closed post-7th"


FOR i = 1 to 19								'formatting the cells'
	objExcel.Cells(2, i).Font.Bold = True	'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 	'formatting as text
	objExcel.Columns(i).AutoFit()			'sizing the columns'
NEXT

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'establishing the row to start searching in the Excel spreadsheet
excel_row = 3

For each worker in worker_array
	back_to_self
	EMWriteScreen CM_mo, 20, 43				'
	EMWriteScreen CM_yr, 20, 46
	Call navigate_to_MAXIS_screen("REPT", "MFCM")			'navigates to MFCM in the current footer month/year'
	EMWriteScreen worker, 21, 13
	transmit

	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason----'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
    has_content_check = trim(has_content_check)
	If has_content_check <> "" then
		Do
			row = 7	'Sets the row to start searching in MAXIS for
			Do
				
				EMReadScreen MAXIS_case_number, 8, row, 6  	'Reading case number
				EMReadScreen client_name, 18, row, 16
				
                'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
				If trim(MAXIS_case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number
					EMReadScreen alt_case_number, 8, row - 1, 6				'then it reads the row above
                    MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'    
                END IF
                
                If trim(MAXIS_case_number) = "" and trim(client_name) = "" then exit do			'Exits do if we reach the end
				
                EMReadScreen vendor_reason, 2, row, 45
                EMReadScreen total_TANF_mo, 2, row, 69
                EMReadScreen ext_60_mo,     2, row, 75
				
				'add case/case information to Excel
        		ObjExcel.Cells(excel_row,  1).Value = worker
        		ObjExcel.Cells(excel_row,  2).Value = trim(MAXIS_case_number)
                ObjExcel.Cells(excel_row,  3).Value = trim(client_name)		
                ObjExcel.Cells(excel_row,  5).Value = trim(vendor_reason)
                ObjExcel.Cells(excel_row, 13).Value = trim(ext_60_mo)
                ObjExcel.Cells(excel_row, 14).Value = trim(total_TANF_mo)	
                	
				excel_row = excel_row + 1	'moving excel row to next row'
				MAXIS_case_number = ""          'Blanking out variable
				row = row + 1	'adding one row to search for in MAXIS
			Loop until row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes back into MFCM and grabs the member # and client name, then cchecks the potentially exempt members for subsidized housing
excel_row = 3           're-establishing the row to start checking the members for
Do
	MAXIS_case_number  = objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    client_name        = objExcel.cells(excel_row, 3).Value	're-establishing the client name to use for the case
    If MAXIS_case_number = "" then exit do						'exits do if the case number is ""
	Call navigate_to_MAXIS_screen("REPT", "MFCM")
	
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete				'row gets deleted since it will get added to the priv case list at end of script 
		IF excel_row = 3 then 
			excel_row = excel_row
		Else 
			excel_row = excel_row - 1
		End if
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	Else 
        EMReadScreen case_content, 7, 8, 7
	    If trim(case_content) = "" then 
	    	'making sure we are getting the right person for cases where there are more than one case. 
        	row = 7
        	Do 
            	EMReadScreen case_name, 18, row, 16
	    		case_name = trim(case_name)
            	If case_name <> client_name then row = row + 1
        	LOOP until case_name = client_name  
	    	EMWriteScreen "x", row, 36		'going into the SANC panel to get case info  
	    Else 
	    	EMWriteScreen "x", 7, 36		'going into the SANC panel to get case info
	    End if 
	
		transmit
	    '----------------------------------------------------------------------------------------------------SANC panel
        EMReadScreen ERRR_panel_check, 4, 2, 52         'Ensuring that there are no errors on the case. 
        If ERRR_panel_check = "ERRR" then transmit      'If they are the client inforamiton will not input.
	    
	    'Reading and inputing information from the SANC panel
	    EMReadScreen memb_number, 2, 4, 12		'reading member number
	    EMReadScreen total_sanc, 2, 17, 43
        total_sanc = trim(total_sanc)
	    
	    EMReadScreen seven_occur, 5, 18, 43
        seven_occur = trim(seven_occur)
	    EMReadScreen post_seven_occur, 5, 19, 43
        post_seven_occur = trim(post_seven_occur)
	    
		If seven_occur <> "" then seven_occur = replace(seven_occur, " ", "/")
	    If post_seven_occur <> "" then post_seven_occur = replace(post_seven_occur, " ", "/")
        
	    ObjExcel.Cells(excel_row,  4).Value = memb_number
		ObjExcel.Cells(excel_row, 15).Value = trim(total_sanc)				
	    ObjExcel.Cells(excel_row, 18).Value = trim(seven_occur)	
	    ObjExcel.Cells(excel_row, 19).Value = trim(post_seven_occur)	
		
		'sanction percentages
		If total_sanc = ""	 then sanc_percent = ""
		If total_sanc =  "1" then sanc_percent = "10"
		If total_sanc =  "2" then sanc_percent = "30"
		If total_sanc =  "3" then sanc_percent = "30"
		If total_sanc =  "4" then sanc_percent = "30"
		If total_sanc =  "5" then sanc_percent = "30"
		If total_sanc =  "6" then sanc_percent = "30"
		If total_sanc =  "7" then sanc_percent = "100"
		If total_sanc =  "8" then sanc_percent = "30"
		If total_sanc =  "9" then sanc_percent = "100"
		If total_sanc = "10" then sanc_percent = "30"
		If total_sanc = "11" then sanc_percent = "100"
		If total_sanc = "12" then sanc_percent = "30"
		If total_sanc = "13" then sanc_percent = "100"
		If total_sanc = "14" then sanc_percent = "30"
		If total_sanc = "15" then sanc_percent = "100"
		If total_sanc = "16" then sanc_percent = "30"
		If total_sanc = "17" then sanc_percent = "100"
		If total_sanc = "18" then sanc_percent = "30"
		If total_sanc = "19" then sanc_percent = "100"
		If total_sanc = "20" then sanc_percent = "30"
		If total_sanc = "21" then sanc_percent = "100"
		If total_sanc = "22" then sanc_percent = "30"
		If total_sanc = "23" then sanc_percent = "100"
		If total_sanc = "24" then sanc_percent = "30"
		If total_sanc = "25" then sanc_percent = "100"
		If total_sanc = "26" then sanc_percent = "30"
		If total_sanc = "27" then sanc_percent = "100"
		If total_sanc = "28" then sanc_percent = "30"
		If total_sanc = "29" then sanc_percent = "100"
		If total_sanc = "30" then sanc_percent = "30"
	
		ObjExcel.Cells(excel_row, 16).Value = trim(sanc_percent)	
		
        'sanction codes for current month - 5 (6 months worth of codes)
        IF CM_mo = "01" then month_col = 10
        IF CM_mo = "02" then month_col = 16
        IF CM_mo = "03" then month_col = 22
        IF CM_mo = "04" then month_col = 28
        IF CM_mo = "05" then month_col = 34
        IF CM_mo = "06" then month_col = 40
        IF CM_mo = "07" then month_col = 46
        IF CM_mo = "08" then month_col = 52
        IF CM_mo = "09" then month_col = 58
        IF CM_mo = "10" then month_col = 64
        IF CM_mo = "11" then month_col = 70
        IF CM_mo = "12" then month_col = 76
        
        If CM_yr = "17" then year_row = 13
        If CM_yr = "18" then year_row = 14
        
		col = month_col 
		row = year_row 
        sanc_count = 0
        sanc_list = ""
    
        Do 
            EMReadScreen sanc_codes, 4, row, col
			'msgbox sanc_codes & vbcr & "row: " & row & vbcr & "col: " & col
            sanc_count = sanc_count + 1
            col = col - 6
            if col = 4 then 
                col = 76
                row = row - 1
            End if 
            If sanc_codes <> "__ _" then sanc_list = sanc_list & sanc_codes & ", "
        Loop until sanc_count = 6
        
        'takes the last comma off of sanc_list 
		sanc_list = trim(sanc_list)
        If right(sanc_list, 1) = "," THEN sanc_list = left(sanc_list, len(sanc_list) - 1) 
        ObjExcel.Cells(excel_row, 17).Value = sanc_list
        'msgbox sanc_list
        '----------------------------------------------------------------------------------------------------ABPS
        Call navigate_to_MAXIS_screen("STAT", "ABPS")
        Call write_value_and_transmit(memb_number, 20, 76)
        
        EMReadScreen support_coop, 1, 4, 73
		If support_coop = "_" then support_coop = ""
        ObjExcel.Cells(excel_row, 6).Value = support_coop
        
        '----------------------------------------------------------------------------------------------------EMPS	
        Call navigate_to_MAXIS_screen("STAT", "EMPS")
        Call write_value_and_transmit(memb_number, 20, 76)

        'Attended Financial orientation code	
        EMReadScreen attended_orient, 1, 5, 65
        If attended_orient = "_" then attended_orient = ""
		
		'Financial orientation date
		EMReadScreen orient_date, 8, 5, 39
        If orient_date = "__ __ __" then 
			orient_date = "" 
		Else 
			orient_date = replace(orient_date, " ", "/")
		End if 
		
		'EMPS Sanc reason
		EMReadScreen EMPS_sanc_reason, 2, 18, 40
		If EMPS_sanc_reason = "__" then EMPS_sanc_reason = ""
		
		'sanction begin date
        EMReadScreen EMPS_sanc_begin_date, 8, 18, 51
        If EMPS_sanc_begin_date = "__ 01 __" then 
			EMPS_sanc_begin_date = "" 
		Else 
			EMPS_sanc_begin_date = replace(EMPS_sanc_begin_date, " ", "/")
		End if
		
		'sanction end date 
		EMReadScreen EMPS_sanc_end_date, 8, 18, 70
        If EMPS_sanc_end_date = "__ 01 __" then 
			EMPS_sanc_end_date = ""
		Else 
			EMPS_sanc_end_date = replace(EMPS_sanc_end_date, " ", "/")
		End if 
        
        ObjExcel.Cells(excel_row, 7).Value = attended_orient
        ObjExcel.Cells(excel_row, 8).Value = orient_date
        ObjExcel.Cells(excel_row, 9).Value = EMPS_sanc_reason
        ObjExcel.Cells(excel_row, 10).Value = EMPS_sanc_begin_date
        ObjExcel.Cells(excel_row, 11).Value = EMPS_sanc_end_date
        
        '----------------------------------------------------------------------------------------------------TIME
        Call navigate_to_MAXIS_screen("STAT", "TIME")
        Call write_value_and_transmit(memb_number, 20, 76)
        
		'time codes for current month and previous month
		IF CM_mo = "01" then month_col = 15
		IF CM_mo = "02" then month_col = 20
		IF CM_mo = "03" then month_col = 25
		IF CM_mo = "04" then month_col = 30
		IF CM_mo = "05" then month_col = 35
		IF CM_mo = "06" then month_col = 40
		IF CM_mo = "07" then month_col = 45
		IF CM_mo = "08" then month_col = 50
		IF CM_mo = "09" then month_col = 55
		IF CM_mo = "10" then month_col = 60
		IF CM_mo = "11" then month_col = 65
		IF CM_mo = "12" then month_col = 70
		
		If CM_yr = "17" then year_row = 13
		If CM_yr = "18" then year_row = 14
		
		col = month_col
		row = year_row    
		time_count = 0
		time_list = ""
	
		Do 
			EMReadScreen time_codes, 2, row, col
			'msgbox time_codes & vbcr & "row: " & row & vbcr & "col: " & col
			time_count = time_count + 1
			col = col - 5
			if col = 10 then 
				col = 70
				row = row - 1
			End if 
			If time_codes <> "__" then time_list = time_list & time_codes & ", "
		Loop until time_count = 2
		
		'takes the last comma off of time_list 
		time_list = trim(time_list)
		If right(time_list, 1) = "," THEN time_list = left(time_list, len(time_list) - 1) 
		ObjExcel.Cells(excel_row, 12).Value = time_list
		
    	'msgbox time_list
        
		excel_row = excel_row + 1
	    STATS_counter = STATS_counter + 1
	End if 
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

IF priv_case_list <> "" then 
	'Creating the list of privileged cases and adding to the spreadsheet
	excel_row = 3				'establishes the row to start writing the PRIV cases to
	objExcel.cells(1, 20).Value = "PRIV cases"
	
	prived_case_array = split(priv_case_list, "|")
	
	FOR EACH MAXIS_case_number in prived_case_array
		If trim(MAXIS_case_number) <> "" then 
			objExcel.cells(excel_row, 20).value = MAXIS_case_number		'inputs cases into Excel
			excel_row = excel_row + 1								'increases the row
		End if 
	NEXT
End if

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
'Query date/time/runtime info
ObjExcel.Cells(1, 21).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, 21).Font.Bold = TRUE
ObjExcel.Cells(2, 21).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, 21).Font.Bold = TRUE
ObjExcel.Cells(3, 21).Value = "Case count:"	'Goes back one, as this is on the next row
objExcel.Cells(3, 21).Font.Bold = TRUE
ObjExcel.Cells(1, 22).Value = now
ObjExcel.Cells(2, 22).Value = timer - query_start_time
ObjExcel.Cells(3, 22).Value = STATS_counter

FOR i = 1 to 22		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'ObjExcel.Cells(1, 1).Value = "CASE INFORMATION"
'ObjExcel.Cells(1, 2).Value = "ABPS PANEL"
'ObjExcel.Cells(1, 3).Value = "TIME PANEL"
'ObjExcel.Cells(1, 4).Value = "SANC PANEL"
'
'FOR i = 1 to 4								'formatting the cells'
'	objExcel.Cells(1, i).Font.Bold = True	'bold font'
'	ObjExcel.columns(i).NumberFormat = "@" 	'formatting as text
'	objExcel.Columns(i).AutoFit()			'sizing the columns'
'NEXT

'ObjExcel.Cells("A1:A5").MergeCells = True

'ObjExcel.Cells("A1:A5").HorizontalAlignment = xlCenter
'ObjExcel.Cells("A6").HorizontalAlignment = xlCenter

'For i = 7 to 11
'	ObjExcel.Cells(1, i).HorizontalAlignment = xlCenter
'Next
'
'For i = 12 to 14
'	ObjExcel.Cells(1, i).HorizontalAlignment = xlCenter
'Next
'
'For i = 15 to 19
'	ObjExcel.Cells(1, i).HorizontalAlignment = xlCenter
'Next

'Range("A1").Borders(xlEdgeBottom).Color = RGB(255, 0, 0)		'HEX Codes for reference 
For i = 1 to 5
	ObjExcel.Columns(i).Interior.Color 	= RGB(208, 206, 206)	' #d0cece
Next 

ObjExcel.Columns(6).Interior.Color 		= RGB(252, 228, 214)	' #fce4d6

For i = 7 to 11
	ObjExcel.Columns(i).Interior.Color 	= RGB(217, 225, 242) 	' #d9e1f2
Next

For i = 12 to 14
	ObjExcel.Columns(i).Interior.Color 	= RGB(255, 242, 204) 	' #fff2cc
Next

For i = 15 to 19
	ObjExcel.Columns(i).Interior.Color 	= RGB(226, 239, 218) 	' #e2efda
Next
	
script_end_procedure("Success! Please review the list generated.")