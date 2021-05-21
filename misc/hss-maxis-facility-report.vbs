'Required for statistical purposes===============================================================================
name_of_script = "MISC - HSS MAXIS FACILITY REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 51                      'manual run time in seconds
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
		FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"   'defaulting everything to Hennepin County Master Functions Libary.
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
call changelog_update("05/21/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""

'----------------------------------Set up code 
MAXIS_footer_month = CM_mo 
MAXIS_footer_year = CM_yr 
case_status = ""            'defaulting case_status to "" to increment later in certain circumsatnces. 
'Excel columns
const HS_status_col     = 16
const vendor_num_col    = 17
const faci_name_col     = 18
const faci_in_col       = 19
const faci_out_col      = 20
const impact_vnd_col    = 21
const exempt_code_col   = 22
const HDL_one_col       = 23
const HDL_two_col       = 24
const HDL_three_col     = 25
const case_status_col   = 26

'User interface dialog - There's just one in this script. 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 90, "HSS MAXIS Facility Report"
  ButtonGroup ButtonPressed
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 65, 50, 15
    CancelButton 420, 65, 50, 15
  EditBox 15, 45, 400, 15, file_selection_path
  Text 15, 20, 455, 20, "This script should be used when adding MAXIS Facility information to an exisiting spreadsheet with an initial data set provided by DHS for the purposes of possible Supplemental Service Rate reductions due to overlapping Housing Stabilization Services (HSS)."
  Text 30, 70, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 80, "Using this script:"
EndDialog

'Display dialog and dialog DO...Loop for mandatory fields and password prompting  
Do 
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case #"
ObjExcel.Cells(1, 3).Value = "Next REVW"
ObjExcel.Cells(1, 4).Value = "Facility Name"
ObjExcel.Cells(1, 5).Value = "GRH Rate"
ObjExcel.Cells(1, 6).Value = "DISA Dates"
ObjExcel.Cells(1, 7).Value = "Certification Dates"
ObjExcel.Cells(1, 8).Value = "GRH Plan Dates"
ObjExcel.Cells(1, 9).Value = "Waiver Type"

ObjExcel.Cells(1, HS_status_col).Value   = date & " MAXIS HS Status"   'col 16
ObjExcel.Cells(1, vendor_num_col).Value  = "Vendor #"                  'col 17
ObjExcel.Cells(1, faci_name_col).Value   = "Facility Name"             'col 18
ObjExcel.Cells(1, faci_in_col).Value     = "Faci In Date"              'col 19
ObjExcel.Cells(1, faci_out_col).Value    = "Faci Out Date"             'col 20
ObjExcel.Cells(1, impact_vnd_col).Value  = "Impacted Vendor?"          'col 21
ObjExcel.Cells(1, exempt_code_col).Value = "VND2 Exemption Code"       'col 22
ObjExcel.Cells(1, HDL_one_col).Value     = "VND2 HDL 1 Code"           'col 23
ObjExcel.Cells(1, HDL_two_col).Value     = "VND2 HDL 2 Code"           'col 24
ObjExcel.Cells(1, HDL_three_col).Value   = "VND2 HDL 3 Code"           'col 25
ObjExcel.Cells(1, case_status_col).Value = "Case Status"               'col 26

FOR i = 1 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

Call check_for_MAXIS(False) 'Ensuring we're actually in MAXIS 
Call MAXIS_footer_month_confirmation(MAXIS_footer_month, MAXIS_footer_year) 'Ensuring we're in the right footer month/year: current footer month/year for this process. 

excel_row = 2 'starting with the 1st non-header row 
Do
    client_PMI = trim(objExcel.cells(excel_row, 1).Value)
    If client_PMI = "" then exit do
    
	MAXIS_case_number      = trim(objExcel.cells(excel_row,  2).Value)
    service_agreement_faci = trim(objExcel.cells(excel_row, 12).Value)
	
    '----------------------------------------------------------------------------------------------------CASE/PERS & PERS Search 
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv) 
    If is_this_priv = True then 
        case_status = "Privileged Case. Unable to access."
    Else 
        Do 
            Call navigate_to_MAXIS_screen("CASE", "PERS")
            row = 10    'staring row for 1st member 
            Do
                EMReadScreen person_PMI, 8, row, 34
                If trim(person_PMI) = client_PMI then 
                    EmReadscreen GRH_status, 1, row, 66
                    If trim(GRH_status) <> "" then
                        EmReadscreen member_number, 2, row, 3 
                        memb_found = True 
                        exit do 
                    Else 
                        'try to match the correct case number in PERS search 
                        back_to_self
                        Call navigate_to_MAXIS_screen("PERS", "    ")
                        Call write_value_and_transmit(client_PMI, 15, 36)
                        EmReadscreen PERS_screen_check, 4, 2, 47
                        If PERS_screen_check = "PERS" then 
                            EmReadscreen PERS_err, 75, 24, 2
                            case_status = trim(PERS_err)
                        Elseif PERS_screen_check <> "PERS" then
                            EmReadscreen match_screen, 4, 2, 51
                            If match_screen = "MTCH" then 
                                EmReadscreen dupe_matches, 11, 9, 7
                                If trim(dupe_matches) <> "" then 
                                    Case_status = "Duplicate exists. Add manually."
                                Else 
                                    'if only one match exists then 
                                    Call write_value_and_transmit("X", 8, 5)
                                    EmReadscreen DSPL_PMI, 8, 5, 44
                                    If trim(DSPL_PMI) = Client_PMI then 
                                        'Read case number after finding HC case 
                                        Call write_value_and_transmit("GR", 7, 22)
                                        EmReadscreen DSPL_case_number, 8, 10, 6
                                        If trim(DSPL_case_number) = "" then 
                                            case_status = "Unable to find HS history for this member."
                                        Else 
                                            MAXIS_case_number = DSPL_case_number
                                        End if         
                                    Else 
                                        case_status = "Unable to find resident by PMI in PERS/DSPL."
                                    End if 
                                End if 
                            End if
                        End if 
                    End if
                Else 
                    row = row + 3			'information is 3 rows apart. Will read for the next member. 
                    If row = 19 then
                        PF8
                        row = 10					'changes MAXIS row if more than one page exists
                    END if
                END if
                EMReadScreen last_PERS_page, 21, 24, 2
            LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
            If memb_found = True then exit do 
        Loop 
    End if 
    
    If trim(member_number) = "" then case_status = "Unable to locate case for member."
    
    If case_status <> "" then     
    
'	call navigate_to_MAXIS_screen("STAT", "FACI")
'  
'	    EMReadScreen FACI_total_check, 1, 2, 78
'	    If FACI_total_check = "0" then 
'	    	current_faci = False 
'			ObjExcel.Cells(excel_row, 4).Value = "Case does not have a FACI panel."	
'	    	case_status = ""
'	    Else 
'	    	row = 14
'	    	Do 
'	    		EMReadScreen date_out, 10, row, 71
'	    		'msgbox "date out: " & date_out 
'	    		If date_out = "__ __ ____" then 
'	 				EMReadScreen date_in, 10, row, 47
'					If date_in <> "__ __ ____" then 
'						current_faci = TRUE
'	    				exit do
'	    			ELSE
'	    				current_faci = False 
'	    				row = row + 1
'	    			End if 
'	    		Else 
'	    			row = row + 1
'	    			'msgbox row
'	    			current_faci = False	
'	    		End if 	
'	    		If row = 19 then 
'	    			transmit
'	    			row = 14
'	    		End if 
'	    		EMReadScreen last_panel, 5, 24, 2
'	    	Loop until last_panel = "ENTER"	'This means that there are no other faci panels
'	    End if 
'		
'	    'GETS FACI NAME AND PUTS IT IN SPREADSHEET, IF CLIENT IS IN FACI.
'	    If current_faci = True then
'	    	EMReadScreen FACI_name, 30, 6, 43
'			EMReadScreen GRH_rate, 1, row, 34	
'	    	ObjExcel.Cells(excel_row, 4).Value = trim(replace(FACI_name, "_", ""))
'			ObjExcel.Cells(excel_row, 5).Value = trim(replace(GRH_rate, "_", ""))
'	    End if 
'		
'	    Call navigate_to_MAXIS_screen("STAT", "DISA")
'		'Reading the disa dates
'		EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
'		EMReadScreen disa_end_date, 10, 6, 69
'		disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
'		disa_end_date = Replace(disa_end_date," ","/")
'		disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
'		If disa_dates = "__/__/____ - __/__/____" then disa_dates = ""
'		ObjExcel.Cells(excel_row, 6).Value = disa_dates
'		
'		EMReadScreen cert_start_date, 10, 9, 47			'reading cert dates
'		EMReadScreen cert_end_date, 10, 9, 69
'		cert_start_date = Replace(cert_start_date," ","/")		'cleans up cert dates
'		cert_end_date = Replace(cert_end_date," ","/")
'		cert_dates = trim(cert_start_date) & " - " & trim(cert_end_date)
'		If cert_dates = "__/__/____ - __/__/____" then cert_dates = ""
'		ObjExcel.Cells(excel_row, 7).Value = cert_dates
'		
'		EMReadScreen GRH_start_date, 10, 9, 47			'reading GRH dates
'		EMReadScreen GRH_end_date, 10, 9, 69
'		GRH_start_date = Replace(GRH_start_date," ","/")		'cleans up GRH dates
'		GRH_end_date = Replace(GRH_end_date," ","/")
'		GRH_dates = trim(GRH_start_date) & " - " & trim(GRH_end_date)
'		If GRH_dates = "__/__/____ - __/__/____" then GRH_dates = ""
'		ObjExcel.Cells(excel_row, 8).Value = GRH_dates
'	    
'	    'checks the waiver type
'	    EMReadScreen DISA_waiver_type, 1, 14, 59
'	    If DISA_waiver_type = "_" then DISA_waiver_type = ""
'	    ObjExcel.Cells(excel_row, 9).Value = DISA_waiver_type
'	End if 
'	
'	excel_row = excel_row + 1 'setting up the script to check the next row.
'LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list
'
''Query date/time/runtime info
'objExcel.Cells(1, 10).Font.Bold = TRUE
'objExcel.Cells(2, 10).Font.Bold = TRUE
'ObjExcel.Cells(1, 10).Value = "Query date and time:"	'Goes back one, as this is on the next row
'ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
'ObjExcel.Cells(1, 11).Value = now
'ObjExcel.Cells(2, 11).Value = timer - query_start_time
'
''formatting the cells
'FOR i = 1 to 11
'	objExcel.Columns(i).AutoFit()				'sizing the columns
'NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")