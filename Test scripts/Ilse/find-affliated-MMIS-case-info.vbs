'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - FIND AFFILIATED MMIS CASE INFO.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 300                      'manual run time in seconds
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
call changelog_update("01/17/2019", "Added function to determine and add recipient's age.", "Ilse Ferris, Hennepin County")
call changelog_update("10/05/2018", "Added identification to MA-DX basis recipients as not coverting to METS.", "Ilse Ferris, Hennepin County")
call changelog_update("09/13/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
'get_county_code

MAXIS_footer_month = CM_mo	'establishing footer month/year
MAXIS_footer_year = CM_yr

file_selection_path = ""

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "MAXIS TO METS Conversion Information"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of conversion cases are provided by the METS team or DHS."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

DIM case_array()
ReDim case_array(21, 0)

'constants for array
const case_number_const     	= 0
const clt_PMI_const 	        = 1
const last_name_const           = 2
const first_name_const          = 3
const client_SSN_const          = 4
const client_age_const          = 5
const HC_status_const           = 6
const revw_date_const           = 7
const waiver_info_const	        = 8
const medicare_info_const       = 9
const first_case_number_const   = 10
const first_type_const 	        = 11
const first_elig_const 	       	= 12
const second_case_number_const  = 13
const second_type_const 	    = 14
const second_elig_const 	  	= 15
const third_case_number_const   = 16
const third_type_const      	= 17
const third_elig_const      	= 18
const case_status               = 19
const basket_number_const       = 20
const rsum_PMI_const            = 21

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do
    'Loops until there are no more cases in the Excel list
    basket_number = objExcel.cells(excel_row, 2).Value   'reading the case number from Excel
    basket_number = Trim(basket_number)

    MAXIS_case_number = objExcel.cells(excel_row, 3).Value   'reading the case number from Excel
    MAXIS_case_number = Trim(MAXIS_case_number)

    Client_PMI = objExcel.cells(excel_row, 4).Value          'reading the PMI from Excel
    Client_PMI = trim(Client_PMI)
    If Client_PMI = "" then exit do

    clients_DOB = objExcel.cells(excel_row, 7).Value          'reading the PMI from Excel
    clients_DOB = trim(clients_DOB)
    Call client_age(clients_DOB, clients_age)

    ReDim Preserve case_array(21, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    case_array(case_number_const,           entry_record) = MAXIS_case_number	'The client information is added to the array'
    case_array(clt_PMI_const,               entry_record) = Client_PMI
    case_array(last_name_const,             entry_record) = ""
    case_array(first_name_const,            entry_record) = ""
    case_array(client_SSN_const,            entry_record) = ""
    case_array(client_age_const,            entry_record) = clients_age
    case_array(HC_status_const,             entry_record) = ""
    case_array(revw_date_const,             entry_record) = ""
    case_array(waiver_info_const,	        entry_record) = ""
    case_array(medicare_info_const,         entry_record) = ""
    case_array(first_case_number_const,   	entry_record) = ""
    case_array(first_type_const, 	        entry_record) = ""
    case_array(first_elig_const, 	        entry_record) = ""
    case_array(second_case_number_const,    entry_record) = ""
    case_array(second_type_const, 	        entry_record) = ""
    case_array(second_elig_const, 	        entry_record) = ""
    case_array(third_case_number_const, 	entry_record) = ""
    case_array(third_type_const,      	    entry_record) = ""
    case_array(third_elig_const,            entry_record) = ""
    case_array(case_status,                 entry_record) = False
    case_array(basket_number_const,         entry_record) =	basket_number
    case_array(rsum_PMI_const,              entry_record) =	""


    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1
Loop
'msgbox entry_record

objExcel.Quit		'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document and don't want the script to be confused
back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

For item = 0 to UBound(case_array, 2)
	MAXIS_case_number = case_array(case_number_const, item)	'Case number is set for each loop as it is used in the FuncLib functions'
    Client_PMI = case_array(clt_PMI_const, item)

    Call navigate_to_MAXIS_screen("CASE", "PERS")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
        case_array(case_status, item) = False
		case_array(clt_PMI_const, item) = MAXIS_case_number & " - PRIV case."
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else
        row = 10
        Do
            EMReadScreen person_PMI, 8, row, 34
            person_PMI = trim(person_PMI)
            IF person_PMI = "" then exit do
            IF Client_PMI = person_PMI then
                EmReadscreen last_name, 15, row, 6
                Case_array(last_name_const, item) = trim(last_name)

                EmReadscreen first_name, 11, row, 22
                Case_array(first_name_const, item) = trim(first_name)

                EmReadscreen Client_SSN, 11, row + 1, 6
                Client_SSN = trim(Client_SSN)
                If Client_SSN = "-  -" then Client_SSN = "" 'blanking out to use the PMI later if no SSN exists
                Case_array(client_SSN_const, item) = replace(Client_SSN, "-", "")

                EMReadScreen HC_status, 1, row, 61      'Reading the HC status for the current month for the member
                Case_array(HC_status_const, item) = HC_status
                Case_array(case_status, item) = True
                'msgbox last_name & vbcr & first_name & vbcr & Client_SSN & vbcr & HC_status
                exit do
            Else
                row = row + 3			'information is 3 rows apart. Will read for the next member.
                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
            END if
            EMReadScreen last_PERS_page, 21, 24, 2
        LOOP until last_PERS_page = "THIS IS THE LAST PAGE"

	    Call navigate_to_MAXIS_screen("STAT", "REVW")      'Reading STAT/REVW information
        EMReadScreen next_review, 8, 9, 70
        If next_review = "__ __ __" then
            Case_array(revw_date_const, item) = ""
        else
            EmReadscreen revw_type, 2, 9, 79
            Case_array(revw_date_const, item) = replace(next_review, " ", "/") & " " & revw_type
        End if
        'msgbox next_review & " " & revw_type
    End if
Next

'-------------------------------------------------------------------------------------------------------------------------------------MMIS portion of the script
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'adding column header information to the Excel list
ObjExcel.Cells(1,  1).Value = "Basket"
ObjExcel.Cells(1,  2).Value = "PMI"
ObjExcel.Cells(1,  3).Value = "RSUM PMI"
ObjExcel.Cells(1,  4).Value = "Last Name"
ObjExcel.Cells(1,  5).Value = "First Name"
ObjExcel.Cells(1,  6).Value = "Client Age"
ObjExcel.Cells(1,  7).Value = "MAXIS HC"
ObjExcel.Cells(1,  8).Value = "Next REVW"
ObjExcel.Cells(1,  9).Value = "Waiver"
ObjExcel.Cells(1, 10).Value = "Medicare"
ObjExcel.Cells(1, 11).Value = "1st case"
ObjExcel.Cells(1, 12).Value = "1st type/prog"
ObjExcel.Cells(1, 13).Value = "1st elig dates"
ObjExcel.Cells(1, 14).Value = "2nd case"
ObjExcel.Cells(1, 15).Value = "2nd type/prog"
ObjExcel.Cells(1, 16).Value = "2nd elig dates"
ObjExcel.Cells(1, 17).Value = "3rd case"
ObjExcel.Cells(1, 18).Value = "3rd type/prog"
ObjExcel.Cells(1, 19).Value = "3rd elig dates"
ObjExcel.Cells(1, 20).Value = "Convert?"

FOR i = 1 to 20 	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
For item = 0 to UBound(case_array, 2)
    Client_SSN = case_array(client_SSN_const, item)
    Client_PMI = case_array(clt_PMI_const, item)
    client_PMI = right("00000000" & client_pmi, 8)

    If case_array(case_status, item) = True then
        'msgbox Client_SSN
        Call MMIS_panel_confirmation("RKEY", 52)
        If Client_SSN = "" then
            Call clear_line_of_text(5, 19)
            EmWriteScreen Client_PMI, 4, 19
        else
            Call clear_line_of_text(4, 19)
            EMWriteScreen Client_SSN, 5, 19
        End if
        Call write_value_and_transmit("I", 2, 19)
        RSEL_row = 7
        Do
            EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52
            EmReadscreen panel_check, 4, 1, 51
            If RSEL_panel_check = "RSEL" then
                EmReadscreen RSEL_SSN, 9, RSEL_row, 48
                If RSEL_SSN = Client_SSN then
                    duplicate_entry = True
                    Call write_value_and_transmit("X", RSEL_row, 2)
                    EmReadscreen panel_check, 4, 1, 51
                else
                    Exit do
                    duplicate_entry = False
                End if
            End if

            If panel_check = "RSUM" then
                'RSUM panel PMI
                EmReadscreen RSUM_PMI, 8, 2, 2
                Case_array(rsum_PMI_const, item) = trim(RSUM_PMI)
                'Waiver info
                EmReadscreen waiver_info, 39, 15, 15
                waiver_info = trim(waiver_info)
                If waiver_info = "BEG DT:          THROUGH DT:" then waiver_info = ""
                Case_array(waiver_info_const, item) = waiver_info
                'Medicare info
                EmReadscreen medicare_info, 69, 21, 10
                medicare_info = trim(medicare_info)
                IF medicare_info = "PART A BEG:          END:          PART B BEG:          END:" then medicare_info = ""
                Case_array(medicare_info_const, item) = medicare_info

                '1st case type/prog/elig/case number
                EmReadscreen first_case_number, 8, 7, 16
                first_case_number = trim(first_case_number)
                If first_case_number <> "" then
                    case_array(first_case_number_const, item) = first_case_number
                    EmReadscreen first_program, 2, 6, 13
                    EmReadscreen first_type, 2, 6, 35
                    If trim(first_program) <> "" then
                        first_elig_type = first_program & "-" & first_type
                        case_array(first_type_const, item) = first_elig_type
                        '1st elig dates
                        EmReadscreen first_elig_start, 8, 7, 35
                        EmReadscreen first_elig_end, 8, 7, 54
                        first_elig_dates = first_elig_start &  " - " & first_elig_end
                        case_array(first_elig_const, item) = first_elig_dates
                    ENd if
                End if

                EmReadscreen second_case_number, 8, 9, 16
                second_case_number = trim(second_case_number)
                If second_case_number <> "" then
                    case_array(second_case_number_const, item) = second_case_number
                    EmReadscreen second_program, 2, 8, 13
                    EmReadscreen second_type, 2, 8, 35
                    If trim(second_program) <> "" then
                        second_elig_type = second_program & "-" & second_type
                        case_array(second_type_const, item) = second_elig_type
                        '1st elig dates
                        EmReadscreen second_elig_start, 8, 9, 35
                        EmReadscreen second_elig_end, 8, 9, 54
                        second_elig_dates = second_elig_start &  " - " & second_elig_end
                        case_array(second_elig_const, item) = second_elig_dates
                    ENd if
                End if

                EmReadscreen third_case_number, 8, 11, 16
                third_case_number = trim(third_case_number)
                If third_case_number <> "" then
                    case_array(third_case_number_const, item) = third_case_number
                    EmReadscreen third_program, 2, 10, 13
                    EmReadscreen third_type, 2, 10, 35
                    If trim(third_program) <> "" then
                        third_elig_type = third_program & "-" & third_type
                        case_array(third_type_const, item) = third_elig_type
                        '1st elig dates
                        EmReadscreen third_elig_start, 8, 11, 35
                        EmReadscreen third_elig_end, 8, 11, 54
                        third_elig_dates = third_elig_start &  " - " & third_elig_end
                        case_array(third_elig_const, item) = third_elig_dates
                    ENd if

                End if
                'outputting to Excel

                If first_case_number <> "" then
                    objExcel.Cells(excel_row,  1).Value = case_array (basket_number_const,      item)
                    objExcel.Cells(excel_row,  2).Value = case_array (clt_PMI_const,            item)
                    objExcel.Cells(excel_row,  3).Value = case_array (rsum_PMI_const,           item)
                    objExcel.Cells(excel_row,  4).Value = case_array (last_name_const,          item)
                    objExcel.Cells(excel_row,  5).Value = case_array (first_name_const,         item)
                    objExcel.Cells(excel_row,  6).Value = case_array (client_age_const,         item)
                    objExcel.Cells(excel_row,  7).Value = case_array (HC_status_const,          item)
                    objExcel.Cells(excel_row,  8).Value = case_array (revw_date_const,          item)
                    objExcel.Cells(excel_row,  9).Value = case_array (waiver_info_const,	    item)
                    objExcel.Cells(excel_row, 10).Value = case_array (medicare_info_const,      item)
                    objExcel.Cells(excel_row, 11).Value = case_array (first_case_number_const,  item)
                    objExcel.Cells(excel_row, 12).Value = case_array (first_type_const, 	    item)
                    objExcel.Cells(excel_row, 13).Value = case_array (first_elig_const, 	    item)
                    objExcel.Cells(excel_row, 14).Value = case_array (second_case_number_const, item)
                    objExcel.Cells(excel_row, 15).Value = case_array (second_type_const, 	    item)
                    objExcel.Cells(excel_row, 16).Value = case_array (second_elig_const, 	    item)
                    objExcel.Cells(excel_row, 17).Value = case_array (third_case_number_const,  item)
                    objExcel.Cells(excel_row, 18).Value = case_array (third_type_const,      	item)
                    objExcel.Cells(excel_row, 19).Value = case_array (third_elig_const,         item)

                    'conditions for converting a case
                    convert_case = ""

                    If left(first_case_number, 1) = "1" then
                        convert_case = False                                'Mets cases starting with 1
                    elseif left(first_case_number, 1) = "2" then
                        convert_case = False                                 'Mets cases starting with 2
                    elseif left(first_case_number, 1) = "S" then
                        convert_case = False                                'special HC programs
                    elseif left(first_elig_type, 2) = "IM" then
                        convert_case = FALSE                                'IMD cases
                    elseif left(first_elig_type, 2) = "NM" then
                        convert_case = FALSE                                  'EMA cases
                    Elseif right(waiver_info, 2) = "19" then
                        convert_case = False                                'ongoing waiver programs
                    elseif right(medicare_info, 2) = "99" then
                        If first_elig_type = "MA-AA" then
                            convert_case = True                             'parent basis with Medicare open
                        else
                            convert_case = False                                'current and ongoing Medicare coverage
                        end if
                    elseif first_elig_type = "MA-11" then
                        convert_case = FALSE          'Auto newborn
                    elseIf first_elig_type = "MA-PX" then
                        convert_case = FALSE          'Pregnant women
                    elseIf first_elig_type = "MA-14" then
                        convert_case = FALSE          'TYMA
                    elseIf first_elig_type = "MA-15" then
                        convert_case = FALSE          '1619 B - disa
                    elseIf first_elig_type = "MA-25" then
                        convert_case = FALSE          'Foster care
                    elseIf first_elig_type = "MA-2A" then
                        convert_case = FALSE          'presumtive elig parent
                    elseIf first_elig_type = "MA-2C" then
                        convert_case = FALSE          'presumtive elig child
                    elseIf first_elig_type = "MA-BT" then
                        convert_case = FALSE          'TEFRA Blind
                    elseIf first_elig_type = "MA-DC" then
                        convert_case = FALSE          'Disabled child
                    elseIf first_elig_type = "MA-DP" then
                        convert_case = FALSE          'MA-EPD
                    elseIf first_elig_type = "MA-DT" then
                        convert_case = FALSE          'TEFRA disabled child
                    elseIf first_elig_type = "MA-EX" then
                        convert_case = FALSE          'Elderly basis
                    elseIf first_elig_type = "MA-DX" then
                        convert_case = FALSE          'Disabled basis
                    else
                        If first_elig_end = "99/99/99" then
                            convert_case = TRUE                             'Open MA cases in MMIS
                        else
                            convert_case = False                            'Closed MA cases in MMIS
                        End if
                    End if

                    If convert_case = False then objExcel.Cells(excel_row, 20).Value = "No"
                    If convert_case = True then objExcel.Cells(excel_row, 20).Value = "Yes"
                    excel_row = excel_row + 1
                End if

                If duplicate_entry = True then
                    PF3
                    EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52
                    If RSEL_panel_check = "RSEL" then
                        RSEL_row = 8
                        case_array(waiver_info_const,	    item) = ""
                        case_array(medicare_info_const,     item) = ""
                        case_array(first_case_number_const, item) = ""
                        case_array(first_type_const, 	    item) = ""
                        case_array(first_elig_const, 	    item) = ""
                        case_array(second_case_number_const,item) = ""
                        case_array(second_type_const, 	    item) = ""
                        case_array(second_elig_const,       item) = ""
                        case_array(third_case_number_const, item) = ""
                        case_array(third_type_const,        item) = ""
                        case_array(third_elig_const,        item) = ""
                    Else
                        exit do 'No more cases on RSEL
                    end if
                else
                    PF3
                    exit do     'cases that did not have more than one known entry
                End if
            End if
        loop
    else
        'msgbox "error case"
        objExcel.Cells(excel_row,  1).Value = case_array (clt_PMI_const,            item)
        excel_row = excel_row + 1
    End if
Next

FOR i = 1 to 17		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")
