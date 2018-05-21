'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - CARL RENEWAL ASSIGNMENT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 100         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Dialog----------------------------------------------------------------------------------------------------
BeginDialog CARL_selection_dialog, 0, 0, 196, 80, "CARL selection"
  EditBox 115, 10, 75, 15, employee_number
  EditBox 115, 30, 35, 15, MAXIS_footer_month
  EditBox 155, 30, 35, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 85, 55, 50, 15
    CancelButton 140, 55, 50, 15
  Text 5, 15, 110, 10, "Please enter your employee ID#:"
  Text 10, 35, 100, 10, "Select the renewal month/year:"
EndDialog

'The script----------------------------------------------------------------------------------------------------
EMConnect ""		'Connecting to BlueZone
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

employee_number = "143495010"

'The main dialog
Do
	Do
		err_msg = ""
		dialog CARL_selection_dialog
        If ButtonPressed = 0 then StopScript
		If isNumeric(employee_number) = False or len(employee_number) <> 9 then err_msg = err_msg & vbNewLine & "* Please enter a numeric employee number."							
        If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a 2-digit valid footer month."									
        If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a 2-digit valid footer year."									
	  IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine										
	LOOP until err_msg = ""		
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Starting the query start time (for the query runtime at the end)
query_start_time = timer
    
'Excel actions----------------------------------------------------------------------------------------------------
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("T:\Eligibility Support\Restricted\Workforce_Management\CARL\Case assignment template.xlsx")
objExcel.Application.DisplayAlerts = False
objExcel.Application.Visible = True
objExcel.worksheets("All regions").Activate			'Activates the "All regions" worksheet and selects the row to start

excel_row = 2
total_cases = 0

array_of_worker_types = array("ADS", "Adults", "DWP", "Families", "YET", "METS")

For each worker_type in array_of_worker_types 
     
    If worker_type = "ADS" then caseloads_to_search = "X127EJ6,X127FE5,X127EK3,X127EK1,X127EK2,X127EJ7,X127EJ8,X127EJ5" & ",X127EH6,X127EM1,X127FE1,X127FI7,X127FH3,X127F3E,X127F3J,X127F3N,X127FI6" & _
		",X127EK9,X127FH5,X127EK5,X127EN7,X127EK6,X127EK4,X127EN6,X127EL1,X127ER6,X127EP8,X127EQ3,X127FG9,X127FI3" & ",X127EM7,X127FI2,X127FG3,X127EM8,X127EM9,X127EJ4" & _
		",X127EH1,X127EH7,X127EH2,X127EH3,X127FH4,X127FI1" & ",X127EP3,X127EP4, X127EP5,X127EP9"
		
    If worker_type = "Adults" then caseloads_to_search = "X127EJ9,X127EQ8,X127EQ9,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127ER1,X127ER2,X127ER3,X127ER4,X127EG8,X127ER5,X127EG5,X127FH2,X127EHD" & _
		",X127EL8,X127EL9,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127FG5" & ",X127EQ1,X127EF7,X127EN5,X127EQ2,X127EF5,X127EK7,X127EF6,X127EQ5,X127EK8,X127EQ4,X127FH9,X127EG6" & _
	    ",X127ED8,X127EH8,X127EAJ,X127EN1,X127EN2,X127EN3,X127EN4,X127ED6,X127ED7,X127EJ2,X127EJ3,X127FH1,X127FG4,X127F3C,X127F3G,X127F3L,X127EJ1,X127EH9,X127EM2,X127FE6,X127EF8,X127EF9,X127EAK,X127EG9,X127EG0" & _
		",X127EE1,X127FB2,X127EG7,X127ED9,X127EE0,X127EH4,X127EH5,X127F3D,X127FH8" & ",X127EP6,X127EP7,X127EG4,X127FG8"   
		       
    If worker_type = "DWP" then caseloads_to_search = "X127EY8,X127EY9,X127EZ1" & ",X127FE7,X127FE8,X127FE9"
	
    If worker_type = "Families" then caseloads_to_search = "X127FD4,X127FD5,X127EZ5,X127FD8,X127EZ8,X127FH6,X127FD6,X127EZ6,X127EZ7,X127FD9,X127FD7,X127EZ0,X127EDD" & _
		",X127ES4,X127ET2,X127ET3,X127FJ2,X127EX3,X127ES8,X127ET1,X127ES7,X127EM5,X127EM6,X127EZ2,X127EZ9,X127ES5,X127EX2,X127ES6,X127EZ4,X127EZ3,X127ES9,X127EX1,X127FF3,X127EW7,X127EW8,X127EW9" & _ 
		",X127EU5,X127EX7,X127F3Y,X127FA3,X127EU6,X127F3S,X127FJ5,X127EY1,X127EY2,X127F3W,X127FA1,X127EU8,X127F3Q,X127EX9,X127FA4,X127BV1,X127F3T,X127FJ1,X127EU9,X127F3X,X127FA2,X127EU7,X127F3R,X127EX8,X127F3Z,X127FJ3,X127FJ4,X127F3V,X127F3U" & _
		",X127EV1,X127FB9,X127FC1,X127EV5,X127FC2,X127EV2,X127EV4,X127EV3,X127FB8,X127FB7" & ",X127ER8,X127ET4,X127F3B,X127ET6,X127ES1,X127ES3,X127FB6,X127ET8,X127F3H,X127F4E,X127FB4,X127F3A,X127F4C,X127F4F,X127FB5,X127F4D,X127F3M,X127ET7,X127FB3,X127ER9,X127ET5,X127ES2,X127BV3" & _
		",X127ET9,X127EU4,X127EW2,X127EW3,X127FH7,X127EU1,X127EU3,X127BV2,X127EU2"
		
    If worker_type = "YET" then caseloads_to_search = "X127CCR,X127CCA,X127FA5,X127FA6,X127FA7,X127FA8,X127FB1,X127FA9"
    If worker_type = "METS" then caseloads_to_search = "X127EN8,X127EN9" & ",X127F4A,X127F4B" & ",X127EX4,X127EX5,X127FF1,X127FF2" & ",X127EQ6,X127EQ7" & ",X127EP1,X127EP2" & ",X127FE2,X127FE3"
    
	'Gathering the information for the Excel spreadsheet
    basket_number_array = split(caseloads_to_search, ",")
	
	back_to_self
	EMWriteScreen CM_mo, 20, 43
	EMWriteScreen CM_yr, 20, 46
    
    For each basket in basket_number_array
    	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
		EMWriteScreen "REPT", 16, 43
		EMWriteScreen "ACTV", 21, 70
		transmit
    	EMWriteScreen basket, 21, 13
    	transmit
    	
		'establishing the renewal date
		renewal_date = MAXIS_footer_month & "/" & MAXIS_footer_year
		
    	'Skips workers with no info
    	EMReadScreen has_content_check, 1, 7, 8
    	If has_content_check <> " " then
    		'Grabbing each case number on screen
    		Do
    			'Set variable for next do...loop
    			MAXIS_row = 7
    
    			'Checking for the last page of cases.
    			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
    			Do
    				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12		 'Reading case number
    				MAXIS_case_number = trim(MAXIS_case_number)
                    EMReadScreen client_name, 21, MAXIS_row, 21		         'Reading client name
    				EMReadScreen review_month, 2, MAXIS_row, 42		         'Reading review month
                    EMReadScreen review_year, 2, MAXIS_row, 48		         'Reading review year
                    EMReadScreen cash_status, 1, MAXIS_row, 54		         'Reading cash status
    				EMReadScreen SNAP_status, 1, MAXIS_row, 61		         'Reading SNAP status
    				EMReadScreen HC_status, 1, MAXIS_row, 64		         'Reading HC status
    				EMReadScreen GRH_status, 1, MAXIS_row, 70		         'Reading GRH status
    
                    review_date = review_month & "/" & review_year
                    
    				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
    				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
    				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
                    If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
    	
                    'Formatting the client name for the spreadsheet
                    client_name = trim(client_name)                     'trimming the client name
                    if instr(client_name, ",") then    'Most cases have both last name and 1st name. This seperates the two names
                        length = len(client_name)                           'establishing the length of the variable   
                        position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
                        last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
                        first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
                    Else                                'In cases where the last name takes up the entire space, then the client name becomes the last name
                        first_name = ""
                        last_name = client_name
                    END IF 
                    if instr(first_name, " ") then   'If there is a middle initial in the first name, then it removes it
                        length = len(first_name)                        'trimming the 1st name
                        position = InStr(first_name, " ")               'establishing the length of the variable
                        first_name = Left(first_name, position-1)       'trims the middle initial off of the first name
                    End if
                
                    'Cleaning up info for spreadsheet. If Inactive, then will show up as blank
                    IF cash_status = "I" then CASH_status = ""
                    IF SNAP_status = "I" then SNAP_status = ""
                    IF HC_status =   "I" then HC_status   = ""
                    IF GRH_status =  "I" then GRH_status  = ""
                        
                    'Adding CARL ID's instead of the population name 
                    If worker_type = "ADS"      then worker_type = "72"
                    If worker_type = "Adults"   then worker_type = "1"     
                    If worker_type = "DWP"      then worker_type = "70" 
                    If worker_type = "Families" then worker_type = "74"
                    If worker_type = "YET"      then worker_type = "75"
                    If worker_type = "METS"     then worker_type = "87"
                
					'REGION ID Key---------------------------------------
					'0: YET only 
					'1: Central/NE
					'2: North Mpls
					'3: Northwest
					'4: South Mpls
					'5: South Suburban
					'6: West
                    'region ID assignments for basket numbers
                    'ADS----------------------------------------------------------------------------------------------------
                    If basket = "X127EJ6" then region_ID = "1"		'1: Central/NE
                    If basket = "X127FE5" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EK3" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EK1" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EK2" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EJ7" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EJ8" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EJ5" then region_ID = "1"		'1: Central/NE
					
                    If basket = "X127EH6" then region_ID = "2"		'2: North Mpls	
                    If basket = "X127EM1" then region_ID = "2"		'2: North Mpls
                    If basket = "X127FE1" then region_ID = "2"		'2: North Mpls
                    If basket = "X127FI7" then region_ID = "2"		'2: North Mpls
                    If basket = "X127FH3" then region_ID = "2"		'2: North Mpls
                    If basket = "X127F3E" then region_ID = "2"		'2: North Mpls
                    If basket = "X127F3J" then region_ID = "2"		'2: North Mpls
                    If basket = "X127F3N" then region_ID = "2"		'2: North Mpls
                    If basket = "X127FI6" then region_ID = "2"		'2: North Mpls
					
                    If basket = "X127EK9" then region_ID = "3"		'3: Northwest
                    If basket = "X127FH5" then region_ID = "3"		'3: Northwest
                    If basket = "X127EK5" then region_ID = "3"		'3: Northwest
					If basket = "X127EN7" then region_ID = "3"		'3: Northwest
                    If basket = "X127EK6" then region_ID = "3"		'3: Northwest
                    If basket = "X127EK4" then region_ID = "3"		'3: Northwest
                    If basket = "X127EN6" then region_ID = "3"		'3: Northwest
                    If basket = "X127EL1" then region_ID = "3"		'3: Northwest
                    If basket = "X127ER6" then region_ID = "3"		'3: Northwest    
                    If basket = "X127EP8" then region_ID = "3"		'3: Northwest
                    If basket = "X127EQ3" then region_ID = "3"		'3: Northwest
                    If basket = "X127FG9" then region_ID = "3"		'3: Northwest
                    If basket = "X127FI3" then region_ID = "3"		'3: Northwest
					
                    If basket = "X127EM7" then region_ID = "4"		'4: South Mpls
                    If basket = "X127FI2" then region_ID = "4"		'4: South Mpls
                    If basket = "X127FG3" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EM8" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EM9" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EJ4" then region_ID = "4"		'4: South Mpls
					
                    If basket = "X127EH1" then region_ID = "5"		'5: South Suburban
                    If basket = "X127EH7" then region_ID = "5"		'5: South Suburban
                    If basket = "X127EH2" then region_ID = "5"		'5: South Suburban
                    If basket = "X127EH3" then region_ID = "5"		'5: South Suburban
                    If basket = "X127FH4" then region_ID = "5"		'5: South Suburban
                    If basket = "X127FI1" then region_ID = "5"		'5: South Suburban
					
                    If basket = "X127EP3" then region_ID = "6"		'6: West
                    If basket = "X127EP4" then region_ID = "6"		'6: West
                    If basket = "X127EP5" then region_ID = "6"		'6: West
                    If basket = "X127EP9" then region_ID = "6"		'6: West
                    'adults----------------------------------------------------------------------------------------------------         
                    If basket = "X127EJ9" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EQ8" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EQ9" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE2" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE3" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE4" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE5" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE6" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EE7" then region_ID = "1"		'1: Central/NE
                    If basket = "X127ER1" then region_ID = "1"		'1: Central/NE
                    If basket = "X127ER2" then region_ID = "1"		'1: Central/NE
                    If basket = "X127ER3" then region_ID = "1"		'1: Central/NE
                    If basket = "X127ER4" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EG8" then region_ID = "1"		'1: Central/NE
                    If basket = "X127ER5" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EG5" then region_ID = "1"		'1: Central/NE
                    If basket = "X127FH2" then region_ID = "1"		'1: Central/NE
                    If basket = "X127EHD" then region_ID = "1"		'1: Central/NE
						
                    If basket = "X127EL8" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL9" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL2" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL3" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL4" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL5" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL6" then region_ID = "2"		'2: North Mpls
                    If basket = "X127EL7" then region_ID = "2"		'2: North Mpls
                    If basket = "X127FG5" then region_ID = "2"		'2: North Mpls
					
                    If basket = "X127EQ1" then region_ID = "3"		'3: Northwest
                    If basket = "X127EF7" then region_ID = "3"		'3: Northwest
                    If basket = "X127EN5" then region_ID = "3"		'3: Northwest
                    If basket = "X127EQ2" then region_ID = "3"		'3: Northwest
                    If basket = "X127EF5" then region_ID = "3"		'3: Northwest
                    If basket = "X127EK7" then region_ID = "3"		'3: Northwest
                    If basket = "X127EF6" then region_ID = "3"		'3: Northwest
                    If basket = "X127EQ5" then region_ID = "3"		'3: Northwest
                    If basket = "X127EK8" then region_ID = "3"		'3: Northwest
                    If basket = "X127EQ4" then region_ID = "3"		'3: Northwest
                    If basket = "X127FH9" then region_ID = "3"		'3: Northwest
                    If basket = "X127EG6" then region_ID = "3"		'3: Northwest
					
                    If basket = "X127ED8" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EH8" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EAJ" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EN1" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EN2" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EN3" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EN4" then region_ID = "4"		'4: South Mpls
                    If basket = "X127ED6" then region_ID = "4"		'4: South Mpls
                    If basket = "X127ED7" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EJ2" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EJ3" then region_ID = "4"		'4: South Mpls
                    If basket = "X127FH1" then region_ID = "4"		'4: South Mpls
                    If basket = "X127FG4" then region_ID = "4"		'4: South Mpls
                    If basket = "X127F3C" then region_ID = "4"		'4: South Mpls
                    If basket = "X127F3G" then region_ID = "4"		'4: South Mpls
                    If basket = "X127F3L" then region_ID = "4"		'4: South Mpls
                    If basket = "X127EJ1" then region_ID = "4"		'4: South Mpls
					If basket = "X127EH9" then region_ID = "4"		'4: South Mpls
					If basket = "X127EM2" then region_ID = "4"		'4: South Mpls
					If basket = "X127FE6" then region_ID = "4"		'4: South Mpls
					'1800 has been assigned to South Mpls
					If basket = "X127EF8" then region_id = "4"		'4: South Mpls
					If basket = "X127EF9" then region_id = "4"		'4: South Mpls
					If basket = "X127EAK" then region_id = "4"		'4: South Mpls
					If basket = "X127EG9" then region_id = "4"		'4: South Mpls
					If basket = "X127EG0" then region_id = "4"		'4: South Mpls
					
                    If basket = "X127EE1" then region_id = "5"		'5: South Suburban
                    If basket = "X127FB2" then region_id = "5"		'5: South Suburban
                    If basket = "X127EG7" then region_id = "5"		'5: South Suburban
                    If basket = "X127ED9" then region_id = "5"		'5: South Suburban
                    If basket = "X127EE0" then region_id = "5"		'5: South Suburban
                    If basket = "X127EH4" then region_id = "5"		'5: South Suburban
                    If basket = "X127EH5" then region_id = "5"		'5: South Suburban
                    If basket = "X127F3D" then region_id = "5"		'5: South Suburban
                    If basket = "X127FH8" then region_id = "5"		'5: South Suburban
					
                    If basket = "X127EP6" then region_id = "6"		'6: West
                    If basket = "X127EP7" then region_id = "6"		'6: West
                    If basket = "X127EG4" then region_id = "6"		'6: West
                    If basket = "X127FG8" then region_id = "6"		'6: West
                    'DWP----------------------------------------------------------------------------------------------------
                    If basket = "X127EY8" then region_id = "2"		'2: North Mpls		
                    If basket = "X127EY9" then region_id = "2"		'2: North Mpls
                    If basket = "X127EZ1" then region_id = "2"		'2: North Mpls
					
                    If basket = "X127FE7" then region_id = "4"		'4: South Mpls		
                    If basket = "X127FE8" then region_id = "4"		'4: South Mpls
                    If basket = "X127FE9" then region_id = "4"		'4: South Mpls
                    'Families----------------------------------------------------------------------------------------------------
                    If basket = "X127FD4" then region_id = "1"		'1: Central/NE
                    If basket = "X127FD5" then region_id = "1"		'1: Central/NE
                    If basket = "X127EZ5" then region_id = "1"		'1: Central/NE
                    If basket = "X127FD8" then region_id = "1"		'1: Central/NE
                    If basket = "X127EZ8" then region_id = "1"		'1: Central/NE
                    If basket = "X127FH6" then region_id = "1"		'1: Central/NE
                    If basket = "X127FD6" then region_id = "1"		'1: Central/NE
                    If basket = "X127EZ6" then region_id = "1"		'1: Central/NE
                    If basket = "X127EZ7" then region_id = "1"		'1: Central/NE
                    If basket = "X127FD9" then region_id = "1"		'1: Central/NE
                    If basket = "X127FD7" then region_id = "1"		'1: Central/NE
                    If basket = "X127EZ0" then region_id = "1"		'1: Central/NE
                    If basket = "X127EDD" then region_id = "1"		'1: Central/NE
					
                    If basket = "X127ES4" then region_id = "2"		'2: North Mpls
                    If basket = "X127ET2" then region_id = "2"		'2: North Mpls
                    If basket = "X127ET3" then region_id = "2"		'2: North Mpls
                    If basket = "X127FJ2" then region_id = "2"		'2: North Mpls
                    If basket = "X127EX3" then region_id = "2"		'2: North Mpls
                    If basket = "X127ES8" then region_id = "2"		'2: North Mpls
                    If basket = "X127ET1" then region_id = "2"		'2: North Mpls
                    If basket = "X127ES7" then region_id = "2"		'2: North Mpls
                    If basket = "X127EM5" then region_id = "2"		'2: North Mpls
                    If basket = "X127EM6" then region_id = "2"		'2: North Mpls
                    If basket = "X127EZ2" then region_id = "2"		'2: North Mpls
                    If basket = "X127EZ9" then region_id = "2"		'2: North Mpls
                    If basket = "X127ES5" then region_id = "2"		'2: North Mpls
                    If basket = "X127EX2" then region_id = "2"		'2: North Mpls
                    If basket = "X127ES6" then region_id = "2"		'2: North Mpls
                    If basket = "X127EZ4" then region_id = "2"		'2: North Mpls
                    If basket = "X127EZ3" then region_id = "2"		'2: North Mpls
                    If basket = "X127ES9" then region_id = "2"		'2: North Mpls
                    If basket = "X127EX1" then region_id = "2"		'2: North Mpls
                    If basket = "X127FF3" then region_id = "2"		'2: North Mpls
                    If basket = "X127EW7" then region_id = "2"		'2: North Mpls
                    If basket = "X127EW8" then region_id = "2"		'2: North Mpls
                    If basket = "X127EW9" then region_id = "2"		'2: North Mpls
					
                    If basket = "X127EU5" then region_id = "3"		'3: Northwest
                    If basket = "X127EX7" then region_id = "3"		'3: Northwest
                    If basket = "X127F3Y" then region_id = "3"		'3: Northwest
                    If basket = "X127FA3" then region_id = "3"		'3: Northwest
                    If basket = "X127EU6" then region_id = "3"		'3: Northwest
                    If basket = "X127F3S" then region_id = "3"		'3: Northwest
                    If basket = "X127FJ5" then region_id = "3"		'3: Northwest
                    If basket = "X127EY1" then region_id = "3"		'3: Northwest
                    If basket = "X127EY2" then region_id = "3"		'3: Northwest
                    If basket = "X127F3W" then region_id = "3"		'3: Northwest
                    If basket = "X127FA1" then region_id = "3"		'3: Northwest
                    If basket = "X127EU8" then region_id = "3"		'3: Northwest
                    If basket = "X127F3Q" then region_id = "3"		'3: Northwest
                    If basket = "X127EX9" then region_id = "3"		'3: Northwest
                    If basket = "X127FA4" then region_id = "3"		'3: Northwest
                    If basket = "X127BV1" then region_id = "3"		'3: Northwest
                    If basket = "X127F3T" then region_id = "3"		'3: Northwest
                    If basket = "X127FJ1" then region_id = "3"		'3: Northwest
                    If basket = "X127EU9" then region_id = "3"		'3: Northwest
                    If basket = "X127F3X" then region_id = "3"		'3: Northwest
                    If basket = "X127FA2" then region_id = "3"		'3: Northwest
                    If basket = "X127EU7" then region_id = "3"		'3: Northwest
                    If basket = "X127F3R" then region_id = "3"		'3: Northwest
                    If basket = "X127EX8" then region_id = "3"		'3: Northwest
                    If basket = "X127F3Z" then region_id = "3"		'3: Northwest
                    If basket = "X127FJ3" then region_id = "3"		'3: Northwest
                    If basket = "X127FJ4" then region_id = "3"		'3: Northwest
                    If basket = "X127F3V" then region_id = "3"		'3: Northwest
                    If basket = "X127F3U" then region_id = "3"		'3: Northwest
					
                    If basket = "X127EV1" then region_id = "4"		'4: South Mpls
                    If basket = "X127FB9" then region_id = "4"		'4: South Mpls
                    If basket = "X127FC1" then region_id = "4"		'4: South Mpls
                    If basket = "X127EV5" then region_id = "4"		'4: South Mpls
                    If basket = "X127FC2" then region_id = "4"		'4: South Mpls
                    If basket = "X127EV2" then region_id = "4"		'4: South Mpls
                    If basket = "X127EV4" then region_id = "4"		'4: South Mpls
					If basket = "X127EV3" then region_id = "4"		'4: South Mpls
                    If basket = "X127FB8" then region_id = "4"		'4: South Mpls
                    If basket = "X127FB7" then region_id = "4"		'4: South Mpls
					
                    If basket = "X127ER8" then region_id = "5"		'5: South Suburban
                    If basket = "X127ET4" then region_id = "5"		'5: South Suburban
                    If basket = "X127F3B" then region_id = "5"		'5: South Suburban
                    If basket = "X127ET6" then region_id = "5"		'5: South Suburban
                    If basket = "X127ES1" then region_id = "5"		'5: South Suburban
                    If basket = "X127ES3" then region_id = "5"		'5: South Suburban
                    If basket = "X127FB6" then region_id = "5"		'5: South Suburban
                    If basket = "X127ET8" then region_id = "5"		'5: South Suburban
                    If basket = "X127F3H" then region_id = "5"		'5: South Suburban
                    If basket = "X127F4E" then region_id = "5"		'5: South Suburban
                    If basket = "X127FB4" then region_id = "5"		'5: South Suburban
                    If basket = "X127F3A" then region_id = "5"		'5: South Suburban
                    If basket = "X127F4C" then region_id = "5"		'5: South Suburban
                    If basket = "X127F4F" then region_id = "5"		'5: South Suburban
                    If basket = "X127FB5" then region_id = "5"		'5: South Suburban
                    If basket = "X127F4D" then region_id = "5"		'5: South Suburban
                    If basket = "X127F3M" then region_id = "5"		'5: South Suburban
                    If basket = "X127ET7" then region_id = "5"		'5: South Suburban
                    If basket = "X127FB3" then region_id = "5"		'5: South Suburban
                    If basket = "X127ER9" then region_id = "5"		'5: South Suburban
                    If basket = "X127ET5" then region_id = "5"		'5: South Suburban
                    If basket = "X127ES2" then region_id = "5"		'5: South Suburban
                    If basket = "X127BV3" then region_id = "5"		'5: South Suburban
					
                    If basket = "X127ET9" then region_id = "6"		'6: West
                    If basket = "X127EU4" then region_id = "6"		'6: West
                    If basket = "X127EW2" then region_id = "6"		'6: West
                    If basket = "X127EW3" then region_id = "6"		'6: West
                    If basket = "X127FH7" then region_id = "6"		'6: West
                    If basket = "X127EU1" then region_id = "6"		'6: West
                    If basket = "X127EU3" then region_id = "6"		'6: West
                    If basket = "X127BV2" then region_id = "6"		'6: West
                    If basket = "X127EU2" then region_id = "6"		'6: West
                    'YET----------------------------------------------------------------------------------------------------
                    If basket = "X127CCR" then region_id = "0"		'0: YET only 	
                    If basket = "X127CCA" then region_id = "0"		'0: YET only 
                    If basket = "X127FA5" then region_id = "0"		'0: YET only 
                    If basket = "X127FA6" then region_id = "0"		'0: YET only 
                    If basket = "X127FA7" then region_id = "0"		'0: YET only 
                    If basket = "X127FA8" then region_id = "0"		'0: YET only 
                    If basket = "X127FB1" then region_id = "0"		'0: YET only 
                    If basket = "X127FA9" then region_id = "0"		'0: YET only 
                    'METS----------------------------------------------------------------------------------------------------
                    If basket = "X127EN8" then region_id = "1"		'1: Central/NE
                    If basket = "X127EN9" then region_id = "1"		'1: Central/NE
					
                    If basket = "X127F4A" then region_id = "2"		'2: North Mpls
                    If basket = "X127F4B" then region_id = "2"		'2: North Mpls
					
                    If basket = "X127EX4" then region_id = "3"		'3: Northwest
                    If basket = "X127EX5" then region_id = "3"		'3: Northwest
                    If basket = "X127FF1" then region_id = "3"		'3: Northwest
                    If basket = "X127FF2" then region_id = "3"		'3: Northwest
					
                    If basket = "X127EQ6" then region_id = "4"		'4: South Mpls
                    If basket = "X127EQ7" then region_id = "4"		'4: South Mpls
					
                    If basket = "X127EP1" then region_id = "5"		'5: South Suburban
                    If basket = "X127EP2" then region_id = "5"		'5: South Suburban
					
                    If basket = "X127FE2" then region_id = "6"		'6: West
                    If basket = "X127FE3" then region_id = "6"		'6: West
                    'End of region ID assignments----------------------------------------------------------------------------------------------------
                    If renewal_date = review_date then 
                        'adding information to the 'script  info' spreadsheet
                        objExcel.Cells(excel_row, 1).Value = region_ID      'number assigned to each region for CARL to read
                        objExcel.Cells(excel_row, 2).Value = "12"           'This is the CARL ID for 'EWS'
                        objExcel.Cells(excel_row, 3).Value = worker_type
                        objExcel.Cells(excel_row, 4).Value = "211"          'This is the CARL ID for 'Work Structure Reviews/Re-certifications'
                        objExcel.Cells(excel_row, 5).Value = MAXIS_case_number
                        objExcel.Cells(excel_row, 6).Value = basket
                        objExcel.Cells(excel_row, 7).Value = first_name
                        objExcel.Cells(excel_row, 8).Value = last_name
                        objExcel.Cells(excel_row, 9).Value = cash_status
                        objExcel.Cells(excel_row, 10).Value = SNAP_status
                        objExcel.Cells(excel_row, 11).Value = HC_status
                        objExcel.Cells(excel_row, 12).Value = GRH_status
    				    objExcel.Cells(excel_row, 13).Value = employee_number
    				    objExcel.Cells(excel_row, 14).Value = date
                        excel_row = excel_row + 1
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
                    End if 
    				total_cases = total_cases + 1
    				MAXIS_row = MAXIS_row + 1
    				add_case_info_to_Excel = ""	'Blanking out variable
    				MAXIS_case_number = ""			'Blanking out variable
    				
    			Loop until MAXIS_row = 19
    			PF8
    		Loop until last_page_check = "THIS IS THE LAST PAGE"
    	End if
        'formatting the cells 
        FOR i = 1 to 13		
        	objExcel.Columns(i).AutoFit()		'sizing the columns
        NEXT
    next
next

STATS_counter = STATS_counter - 1           'starts with one count, so one count needs to be removed.

renewal_date = MAXIS_footer_month & "-" & MAXIS_footer_year

'Adding script inforamtional data AND saving and closing actions----------------------------------------------------------------------------------------------------
objExcel.worksheets("Script_info").Activate         'Activates the informational workesheet (2nd worksheet)
'adding information to the 'script  info' spreadsheet
objExcel.Cells(1, 2).Value = employee_number
objExcel.Cells(2, 2).Value = timer - query_start_time
objExcel.Cells(3, 2).Value = total_cases
objExcel.Cells(4, 2).Value = renewal_date
objExcel.Cells(5, 2).Value = date

'Saves and closes the Excel workbook
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\Workforce_Management\CARL\Renewals.xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

script_end_procedure("Success! The Excel spreahsheet with the case assignment inforamtion has been saved.")