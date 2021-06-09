'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - LANGUAGE STATS.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 20         	'manual run time in seconds
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'The script----------------------------------------------------------------------------------------------------
EMConnect ""		'Connecting to BlueZone

DIM language_info_array()
ReDIM language_info_array(25, 0)

 Const region_ID 	= 1
 Const Amharic  	= 2
 Const Arabic 		= 3
 Const ASL 			= 4
 Const Burmese 		= 5
 Const Cantonese	= 6
 Const English 		= 7
 Const French 		= 8
 Const Hmong 		= 9
 Const Khmer 		= 10
 Const Korean 		= 11
 Const Karen 		= 12
 Const Laotian 		= 13
 Const Mandarin 	= 14
 Const Oromo 		= 15
 Const Russian 		= 16
 Const Serbo 		= 17
 Const Somali 		= 18
 Const Spanish 		= 19
 Const Swahili 		= 20
 Const Tigrinya 	= 21
 Const Vietnamese 	= 22
 Const Yoruba 		= 23
 Const Unknown 		= 24
 Const Other 		= 25

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 296, 90, "Regional Language Statistics"
  ButtonGroup ButtonPressed
    OkButton 185, 70, 50, 15
    CancelButton 240, 70, 50, 15
  Text 15, 25, 265, 20, "This script will gather language stats for each region. It will be going through every active case in the county, so it can take upwards of 8 hours or more to run."
  GroupBox 10, 10, 280, 55, "About this script:"
  Text 35, 50, 225, 10, " Please shut down your VGO (not pause it), and press OK to continue."
EndDialog
'The main dialog
Do
	Do
		dialog Dialog1
        cancel_without_confirmation
	LOOP until ButtonPressed = -1					'This is the OK button
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)
back_to_self

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'The beginning of the arrays----------------------------------------------------------------------------------------------------
region_types_array = array("Central NE", "North Mpls", "Northwest", "South Mpls", "South Suburban", "West")

region = -1			'establishes the value of the region for the array. Since array values start at 0, region starts at -1

For each region_name in region_types_array
	region = region + 1
	case_numbers_array = ""

	'Creating new variables for cases to search based on the population type and region
	If region_name = "Central NE" then caseloads_to_search = "X127EJ6,X127FE5,X127EK3,X127EK1,X127EK2,X127EJ7,X127EJ8,X127EJ5,X127EJ9,X127EQ8,X127EQ9,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127ER1,X127ER2,X127ER3,X127ER4,X127EG8,X127ER5,X127EG5,X127FH2,X127EHD,X127FD4,X127FD5,X127EZ5,X127FD8,X127EZ8,X127FH6,X127FD6,X127EZ6,X127EZ7,X127FD9,X127FD7,X127EZ0,X127EDD,X127EN8,X127EN9"
    If region_name = "North Mpls" then caseloads_to_search = "X127EH6,X127EM1,X127FE1,X127FI7,X127FH3,X127F3E,X127F3J,X127F3N,X127FI6,X127EL8,X127EL9,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127FG5,X127EY8,X127EY9,X127EZ1,X127ES4,X127ET2,X127ET3,X127FJ2,X127EX3,X127ES8,X127ET1,X127ES7,X127EM5,X127EM6,X127EZ2,X127EZ9,X127ES5,X127EX2,X127ES6,X127EZ4,X127EZ3,X127ES9,X127EX1,X127FF3,X127EW7,X127EW8,X127EW9,X127F4A,X127F4B"
    If region_name = "Northwest" then caseloads_to_search = "X127EK9,X127FH5,X127EK5,X127EN7,X127EK6,X127EK4,X127EN6,X127EL1,X127ER6,X127EP8,X127EQ3,X127FG9,X127FI3,X127EQ1,X127EF7,X127EN5,X127EQ2,X127EF5,X127EK7,X127EF6,X127EQ5,X127EK8,X127EQ4,X127FH9,X127EG6,X127EU5,X127EX7,X127F3Y,X127FA3,X127EU6,X127F3S,X127FJ5,X127EY1,X127EY2,X127F3W,X127FA1,X127EU8,X127F3Q,X127EX9,X127FA4,X127BV1,X127F3T,X127FJ1,X127EU9,X127F3X,X127FA2,X127EU7,X127F3R,X127EX8,X127F3Z,X127FJ3,X127FJ4,X127F3V,X127F3U,X127EX4,X127EX5,X127FF1,X127FF2"
    If region_name = "South Mpls" then caseloads_to_search = "X127EM7,X127FI2,X127FG3,X127EM8,X127EM9,X127EJ4,X127ED8,X127EH8,X127EAJ,X127EN1,X127EN2,X127EN3,X127EN4,X127ED6,X127ED7,X127EJ2,X127EJ3,X127FH1,X127FG4,X127F3C,X127F3G,X127F3L,X127EJ1,X127EH9,X127EM2,X127FE6,X127EF8,X127EF9,X127EAK,X127EG9,X127EG0,X127FE7,X127FE8,X127FE9,X127EV1,X127FB9,X127FC1,X127EV5,X127FC2,X127EV2,X127EV4,X127EV3,X127FB8,X127FB7,X127EQ6,X127EQ7"
    If region_name = "South Suburban" then caseloads_to_search = "X127EH1,X127EH7,X127EH2,X127EH3,X127FH4,X127FI1,X127EE1,X127FB2,X127EG7,X127ED9,X127EE0,X127EH4,X127EH5,X127F3D,X127FH8,X127ER8,X127ET4,X127F3B,X127ET6,X127ES1,X127ES3,X127FB6,X127ET8,X127F3H,X127F4E,X127FB4,X127F3A,X127F4C,X127F4F,X127FB5,X127F4D,X127F3M,X127ET7,X127FB3,X127ER9,X127ET5,X127ES2,X127BV3,X127EP1,X127EP2"
    If region_name = "West" then caseloads_to_search = "X127EP3,X127EP4,X127EP5,X127EP9,X127EP6,X127EP7,X127EG4,X127FG8,X127ET9,X127EU4,X127EW2,X127EW3,X127FH7,X127EU1,X127EU3,X127BV2,X127EU2,X127FE2,X127FE3"

	'establishing the count at 0 for each language
	Amharic_count		= 0
	Arabic_count 		= 0
	ASL_count 			= 0
	Burmese_count 		= 0
	Cantonese_count 	= 0
	English_count 		= 0
	French_count 		= 0
	Hmong_count 		= 0
	Khmer_count 		= 0
	Korean_count 		= 0
	Karen_count 		= 0
	Laotian_count 		= 0
	Mandarin_count 		= 0
	Oromo_count 		= 0
	Russian_count 		= 0
	Serbo_count		 	= 0
	Somali_count 		= 0
	Spanish_count 		= 0
	Swahili_count 		= 0
	Tigrinya_count 		= 0
	Vietnamese_count 	= 0
	Yoruba_count 		= 0
	Unknown_count 		= 0
	Other_count 		= 0

	'msgbox caseloads_to_search & vbcr & "region name: " & region_name
	'Gathering the information for the Excel spreadsheet
	basket_number_array = split(caseloads_to_search, ",")

	For each basket in basket_number_array
		'msgbox basket

	    Call navigate_to_MAXIS_screen("rept", "actv")
	    EMWriteScreen basket, 21, 13
	    transmit

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
					If MAXIS_case_number = "" then exit do			'Exits do if we reach the end

					'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
					If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
					all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
					If MAXIS_case_number <> "" then case_number_list = case_number_list & MAXIS_case_number & ","

	    			STATS_counter = STATS_counter + 1               'adds one instance to the stats counter
					MAXIS_row = MAXIS_row + 1
					'MAXIS_case_number = ""							'Blanking out variable
	    		Loop until MAXIS_row = 19
	    		PF8
	    	Loop until last_page_check = "THIS IS THE LAST PAGE"
			'MAXIS_case_number = ""							'Blanking out variable
	    End if
	next

	Erase basket_number_array
	caseloads_to_search = ""

	case_number_list = trim(case_number_list)
	If right(case_number_list, 1) = "," then case_number_list = left(case_number_list, len(case_number_list) - 1)
	case_numbers_array = split(case_number_list, ",")

	'msgbox case_number_list
	For each MAXIS_case_number in case_numbers_array
		IF MAXIS_case_number = "" then exit for
		Call navigate_to_MAXIS_screen ("STAT", "MEMB")
		EMReadScreen language_ID, 2, 12, 42
		If isnumeric(language_ID) = true then
			IF language_ID = "09" then Amharic_count = Amharic_count + 1
			If language_ID = "10" then Arabic_count = Arabic_count + 1
			IF language_ID = "08" then ASL_count = ASL_count + 1
			IF language_ID = "14" then Burmese_count = Burmese_count + 1
			IF language_ID = "15" then Cantonese_count = Cantonese_count + 1
			IF language_ID = "99" then English_count = English_count + 1
			IF language_ID = "16" then French_count = French_count + 1
			IF language_ID = "02" then Hmong_count = Hmong_count + 1
			IF language_ID = "04" then Khmer_count = Khmer_count + 1
			IF language_ID = "20" then Korean_count = Korean_count + 1
			IF language_ID = "21" then Karen_count = Karen_count + 1
			IF language_ID = "05" then Laotian_count = Laotian_count + 1
			IF language_ID = "17" then Mandarin_count = Mandarin_count + 1
			IF language_ID = "12" then Oromo_count = Oromo_count + 1
			IF language_ID = "06" then Russian_count = Russian_count + 1
			IF language_ID = "11" then Serbo_count = Serbo_count + 1
			IF language_ID = "07" then Somali_count = Somali_count + 1
			IF language_ID = "01" then Spanish_count = Spanish_count + 1
			IF language_ID = "18" then Swahili_count = Swahili_count + 1
			IF language_ID = "13" then Tigrinya_count = Tigrinya_count =  + 1
			IF language_ID = "03" then Vietnamese_count = Vietnamese_count + 1
			IF language_ID = "19" then Yoruba_count = Yoruba_count + 1
			IF language_ID = "97" then Unknown_count = Unknown_count + 1
			IF language_ID = "98" then Other_count = Other_count + 1
		END IF
			'End of region ID assignments----------------------------------------------------------------------------------------------------
	next

	Erase case_numbers_array
	case_number_list = ""

	'Adding information to the array for the region
	Redim Preserve language_info_array(25, region)
	language_info_array(region_ID,  region) = region_name
	language_info_array(Amharic, 	region) = Amharic_count
	language_info_array(Arabic, 	region) = Arabic_count
	language_info_array(ASL, 		region) = ASL_count
	language_info_array(Burmese, 	region) = Burmese_count
	language_info_array(Cantonese,  region) = Cantonese_count
	language_info_array(English, 	region) = English_count
	language_info_array(French, 	region) = French_count
	language_info_array(Hmong, 		region) = Hmong_count
	language_info_array(Khmer, 		region) = Khmer_count
	language_info_array(Korean, 	region) = Korean_count
	language_info_array(Karen, 		region) = Karen_count
	language_info_array(Laotian, 	region) = Laotian_count
	language_info_array(Mandarin, 	region) = Mandarin_count
	language_info_array(Oromo, 		region) = Oromo_count
	language_info_array(Russian, 	region) = Russian_count
	language_info_array(Serbo, 		region) = Serbo_count
	language_info_array(Somali, 	region) = Somali_count
	language_info_array(Spanish, 	region) = Spanish_count
	language_info_array(Swahili, 	region) = Swahili_count
	language_info_array(Tigrinya, 	region) = Tigrinya_count
	language_info_array(Vietnamese,	region) = Vietnamese_count
	language_info_array(Yoruba, 	region) = Yoruba_count
	language_info_array(Unknown, 	region) = Unknown_count
	language_info_array(Other, 		region) = Other_count
next

STATS_counter = STATS_counter - 1           'starts with one count, so one count needs to be removed.

'Adding script inforamtional data AND saving and closing actions----------------------------------------------------------------------------------------------------
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "LANGUAGE"
ObjExcel.Cells(1, 2).Value = "CENTRAL N/E"
ObjExcel.Cells(1, 3).Value = "NORTH"
ObjExcel.Cells(1, 4).Value = "NORTHWEST"
ObjExcel.Cells(1, 5).Value = "SOUTH MPLS"
ObjExcel.Cells(1, 6).Value = "SOUTH SUB"
ObjExcel.Cells(1, 7).Value = "WEST"

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

col_to_use = 1

'All the languages
ObjExcel.Cells( 2, col_to_use).Value ="Amharic"
ObjExcel.Cells( 3, col_to_use).Value ="Arabic"
ObjExcel.Cells( 4, col_to_use).Value ="ASL"
ObjExcel.Cells( 5, col_to_use).Value ="Burmese"
ObjExcel.Cells( 6, col_to_use).Value ="Cantonese"
ObjExcel.Cells( 7, col_to_use).Value ="English"
ObjExcel.Cells( 8, col_to_use).Value ="French"
ObjExcel.Cells( 9, col_to_use).Value ="Hmong"
ObjExcel.Cells(10, col_to_use).Value ="Khmer"
ObjExcel.Cells(11, col_to_use).Value ="Korean"
ObjExcel.Cells(12, col_to_use).Value ="Karen"
ObjExcel.Cells(13, col_to_use).Value ="Laotian"
ObjExcel.Cells(14, col_to_use).Value ="Mandarin"
ObjExcel.Cells(15, col_to_use).Value ="Oromo"
ObjExcel.Cells(16, col_to_use).Value ="Russian"
ObjExcel.Cells(17, col_to_use).Value ="Serbo-Croatian"
ObjExcel.Cells(18, col_to_use).Value ="Somali"
ObjExcel.Cells(19, col_to_use).Value ="Spanish"
ObjExcel.Cells(20, col_to_use).Value ="Swahili"
ObjExcel.Cells(21, col_to_use).Value ="Tigrinya"
ObjExcel.Cells(22, col_to_use).Value ="Vietnamese"
ObjExcel.Cells(23, col_to_use).Value ="Yoruba"
ObjExcel.Cells(24, col_to_use).Value ="Unknown"
ObjExcel.Cells(25, col_to_use).Value ="Other"
ObjExcel.Cells(26, col_to_use).Value ="Totals:"

FOR i = 1 to 26											'formatting the cells'
	objExcel.Cells(i, col_to_use).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()						'sizing the columns'
NEXT

col_to_use = col_to_use + 1

'End of setting up the Excel sheet----------------------------------------------------------------------------------------------------
For i = 0 to Ubound(language_info_array, 2)
	'increased by 1 column for each region
	ObjExcel.Cells(2,  col_to_use).Value = language_info_array (Amharic, 	i)
	ObjExcel.Cells(3,  col_to_use).Value = language_info_array (Arabic,   	i)
	ObjExcel.Cells(4,  col_to_use).Value = language_info_array (ASL, 		i)
	ObjExcel.Cells(5,  col_to_use).Value = language_info_array (Burmese, 	i)
	ObjExcel.Cells(6,  col_to_use).Value = language_info_array (Cantonese, 	i)
	ObjExcel.Cells(7,  col_to_use).Value = language_info_array (English, 	i)
	ObjExcel.Cells(8,  col_to_use).Value = language_info_array (French, 	i)
	ObjExcel.Cells(9,  col_to_use).Value = language_info_array (Hmong, 		i)
	ObjExcel.Cells(10, col_to_use).Value = language_info_array (Khmer, 		i)
	ObjExcel.Cells(11, col_to_use).Value = language_info_array (Korean, 	i)
	ObjExcel.Cells(12, col_to_use).Value = language_info_array (Karen, 		i)
	ObjExcel.Cells(13, col_to_use).Value = language_info_array (Laotian, 	i)
	ObjExcel.Cells(14, col_to_use).Value = language_info_array (Mandarin, 	i)
	ObjExcel.Cells(15, col_to_use).Value = language_info_array (Oromo, 		i)
	ObjExcel.Cells(16, col_to_use).Value = language_info_array (Russian, 	i)
	ObjExcel.Cells(17, col_to_use).Value = language_info_array (Serbo, 		i)
	ObjExcel.Cells(18, col_to_use).Value = language_info_array (Somali, 	i)
	ObjExcel.Cells(19, col_to_use).Value = language_info_array (Spanish, 	i)
	ObjExcel.Cells(20, col_to_use).Value = language_info_array (Swahili, 	i)
	ObjExcel.Cells(21, col_to_use).Value = language_info_array (Tigrinya, 	i)
	ObjExcel.Cells(22, col_to_use).Value = language_info_array (Vietnamese, i)
	ObjExcel.Cells(23, col_to_use).Value = language_info_array (Yoruba, 	i)
	ObjExcel.Cells(24, col_to_use).Value = language_info_array (Unknown, 	i)
	ObjExcel.Cells(25, col_to_use).Value = language_info_array (Other, 		i)
	col_to_use = col_to_use + 1
Next

'Adding up the information in the last row
ObjExcel.Cells(26, 2).Value = "=SUM(B2:B25)"
ObjExcel.Cells(26, 3).Value = "=SUM(C2:C25)"
ObjExcel.Cells(26, 4).Value = "=SUM(D2:D25)"
ObjExcel.Cells(26, 5).Value = "=SUM(E2:E25)"
ObjExcel.Cells(26, 6).Value = "=SUM(F2:F25)"
ObjExcel.Cells(26, 7).Value = "=SUM(G2:G25)"

FOR i = 1 to 26		'formatting the cells'
	objExcel.Cells(i, col_to_use).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

col_to_use = col_to_use + 1
objExcel.Cells(1, col_to_use).Value = "REPORT INFORMATION"
objExcel.Cells(2, col_to_use).Value = "SCRIPT RUNTIME:"
objExcel.Cells(3, col_to_use).Value = "TOTAL CASES:"
objExcel.Cells(4, col_to_use).Value = "REPORT DATE:"

FOR i = 1 to 4									'formatting the cells'
	objExcel.Cells(i, col_to_use).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

col_to_use = col_to_use + 1
objExcel.Cells(2, col_to_use).Value = timer - query_start_time
objExcel.Cells(3, col_to_use).Value = stats_counter
objExcel.Cells(4, col_to_use).Value = date

FOR i = 1 to col_to_use						'formatting the cells'
	objExcel.Columns(i).AutoFit()			'sizing the columns'
NEXT

script_end_procedure("")

'EXTRA code that is not needed now, but may want for later

'If region_list = "All regions" then caseloads_to_search = "X127EJ6,X127FE5,X127EK3,X127EK1,X127EK2,X127EJ7,X127EJ8,X127EJ5,X127EJ9,X127EQ8,X127EQ9,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127ER1,X127ER2,X127ER3,X127ER4,X127EG8,X127ER5,X127EG5,X127FH2,X127EHD,X127FD4,X127FD5,X127EZ5,X127FD8,X127EZ8,X127FH6,X127FD6,X127EZ6,X127EZ7,X127FD9,X127FD7,X127EZ0,X127EDD,X127EN8,X127EN9,X127EH6,X127EM1,X127FE1,X127FI7,X127FH3,X127F3E,X127F3J,X127F3N,X127FI6,X127EL8,X127EL9,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127FG5,X127EY8,X127EY9,X127EZ1,X127ES4,X127ET2,X127ET3,X127FJ2,X127EX3,X127ES8,X127ET1,X127ES7,X127EM5,X127EM6,X127EZ2,X127EZ9,X127ES5,X127EX2,X127ES6,X127EZ4,X127EZ3,X127ES9,X127EX1,X127FF3,X127EW7,X127EW8,X127EW9,X127F4A,X127F4B,X127EK9,X127FH5,X127EK5,X127EN7,X127EK6,X127EK4,X127EN6,X127EL1,X127ER6,X127EP8,X127EQ3,X127FG9,X127FI3,X127EQ1,X127EF7,X127EN5,X127EQ2,X127EF5,X127EK7,X127EF6,X127EQ5,X127EK8,X127EQ4,X127FH9,X127EG6,X127EU5,X127EX7,X127F3Y,X127FA3,X127EU6,X127F3S,X127FJ5,X127EY1,X127EY2,X127F3W,X127FA1,X127EU8,X127F3Q,X127EX9,X127FA4,X127BV1,X127F3T,X127FJ1,X127EU9,X127F3X,X127FA2,X127EU7,X127F3R,X127EX8,X127F3Z,X127FJ3,X127FJ4,X127F3V,X127F3U,X127EX4,X127EX5,X127FF1,X127FF2,X127EM7,X127FI2,X127FG3,X127EM8,X127EM9,X127EJ4,X127ED8,X127EH8,X127EAJ,X127EN1,X127EN2,X127EN3,X127EN4,X127ED6,X127ED7,X127EJ2,X127EJ3,X127FH1,X127FG4,X127F3C,X127F3G,X127F3L,X127EJ1,X127EH9,X127EM2,X127FE6,X127EF8,X127EF9,X127EAK,X127EG9,X127EG0,X127FE7,X127FE8,X127FE9,X127EV1,X127FB9,X127FC1,X127EV5,X127FC2,X127EV2,X127EV4,X127EV3,X127FB8,X127FB7,X127EQ6,X127EQ7,X127EH1,X127EH7,X127EH2,X127EH3,X127FH4,X127FI1,X127EE1,X127FB2,X127EG7,X127ED9,X127EE0,X127EH4,X127EH5,X127F3D,X127FH8,X127ER8,X127ET4,X127F3B,X127ET6,X127ES1,X127ES3,X127FB6,X127ET8,X127F3H,X127F4E,X127FB4,X127F3A,X127F4C,X127F4F,X127FB5,X127F4D,X127F3M,X127ET7,X127FB3,X127ER9,X127ET5,X127ES2,X127BV3,X127EP1,X127EP2,X127EP3,X127EP4, X127EP5,X127EP9,X127EP6,X127EP7,X127EG4,X127FG8,X127ET9,X127EU4,X127EW2,X127EW3,X127FH7,X127EU1,X127EU3,X127BV2,X127EU2,X127FE2,X127FE3"

''msgbox "region name: " & region_name & vbcr & _
'	Amharic_count & vbcr & _
'	Arabic_count & vbcr & _
'	ASL_count & vbcr & _
'	Burmese_count & vbcr & _
'	Cantonese_count & vbcr & _
'	English_count & vbcr & _
'	French_count & vbcr & _
'	Hmong_count & vbcr & _
'	Khmer_count  & vbcr & _
'	Korean_count & vbcr & _
'	Karen_count & vbcr & _
'	Laotian_count & vbcr & _
'	Mandarin_count & vbcr & _
'	Oromo_count & vbcr & _
'	Russian_count & vbcr & _
'	Serbo_count & vbcr & _
'	Somali_count & vbcr & _
'	Spanish_count & vbcr & _
'	Swahili_count & vbcr & _
'	Tigrinya_count & vbcr & _
'	Vietnamese_count & vbcr & _
'	Yoruba_count & vbcr & _
'	Unknown_count & vbcr & _
'	Other_count & vbcr & _
'	"stats counter: " & STATS_counter &  vbcr & _
'	"region: " & region
