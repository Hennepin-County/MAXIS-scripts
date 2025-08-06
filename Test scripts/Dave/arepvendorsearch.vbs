
'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - AVS Panel Report.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

call changelog_update("04/25/2025", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Constants
vendor_col  = 1
case_col   = 2
name_col   = 3
last_issuance_col = 4
arep_name_col = 5
arep_addr_col = 6
arep_phone_col = 7

maxis_footer_month = "07"
maxis_footer_year = "25"
'FILE SELECTION DIALOG
DO
	call file_selection_system_dialog(excel_file_path, ".xlsx")
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
	objExcel.Visible = True
	objExcel.DisplayAlerts = True
	confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
	IF confirm_file = vbCancel THEN
		objWorkbook.Close
		objExcel.Quit
		stopscript
	ELSEIF confirm_file = vbNo THEN
		objWorkbook.Close
		objExcel.Quit
	END IF
LOOP UNTIL confirm_file = vbYes

Call check_for_MAXIS(true)
'Load all the vendor numbers into an array
'Dim vendor_array()
'excel_row = 1
'ObjExcel.Sheets(1).Activate
'Do while ObjExcel.Cells(excel_row, vendor_col).value <> "" 'Loop until the vendor number is blank
'    redim preserve vendor_array(excel_row-1) 'resize the array to hold the vendor numbers
'    vendor_array(excel_row-1) = ObjExcel.Cells(excel_row, vendor_col).value
'    'MsgBox "Vendor array: " & vendor_array(excel_row - 1) & "excel: " & ObjExcel.Cells(excel_row, vendor_col).value
'    excel_row = excel_row + 1
'Loop 
'MsgBox "Vendor array: " & vendor_array(23)
'ObjExcel.Sheets(2).Activate
'Set objDictionary = CreateObject("Scripting.Dictionary") 'this will hold unique names
'MsgBox ubound(vendor_array) 
'excel_row = 2 'start at row 2, row 1 is the header
'For vends = 0 to ubound(vendor_array) 'loop through the vendor numbers
'    If vendor_array(vends) <> "" Then
'    vnd_number = vendor_array(vends) 'get the vendor number from the array
'    objDictionary.RemoveAll 'clear the dictionary for each vendor number
'    Call navigate_to_MAXIS_screen("MONY", "VNDW")
'    Call write_value_and_transmit(vnd_number, 4, 11) 
'    EMREadScreen terminate, 11, 24, 2
'    If terminate = "THIS VENDOR" Then
'        PF3
'    Else
'    page_count = 0
'    DO
'        For vnd_row = 7 to 18
'            EMREadScreen clt_name, 23, vnd_row, 33
'            clt_name = Trim(clt_name)
'            If clt_name <> "" Then
'                If objDictionary.exists(clt_name) Then
'                    objdictionary.item(clt_name) = vnd_row 'if the name already exists, change the row number
'                Else
'                    objDictionary.Add clt_name, vnd_row 'add unique name to dictionary
'                    EMREadScreen last_issuance, 8, vnd_row, 14
'                    Call write_value_and_transmit("I", vnd_row, 3)
'                    EMReadScreen Priv_check, 11, 24, 2
'                    If Priv_check = "YOU ARE NOT" Then
'                        EMREadScreen case_number, 8, 18, 43
'                        case_number = replace(case_number, "_", "")
'                        case_number = Trim(case_number)
'                        EMWRitescreen "3_______", 18, 43
'                        MAXIS_case_number = "3"
'                        Call navigate_to_MAXIS_screen("MONY", "VNDW")
'                        Call write_value_and_transmit(vnd_number, 4, 11) 
'                    Else 
'                        EMReadscreen case_number, 8, 19, 38
'                        PF3
'                    End If
'                    ObjExcel.Sheets(2).Cells(excel_row, vendor_col) = vnd_number
'                    ObjExcel.Sheets(2).Cells(excel_row, case_col) = case_number
'                    ObjExcel.Sheets(2).Cells(excel_row, name_col) = clt_name
'                    ObjExcel.Sheets(2).Cells(excel_row, last_issuance_col) = last_issuance
'                    excel_row = excel_row + 1
'                    
'                    If page_count > 0 Then
'                        For transmits = 1 to page_count
'                            Emsendkey "PF8"
'                        Next
'                    End IF  
'                End If 'End of unique name check
'
'            End If
'        Next 
'        PF8
'        EMREadScreen page_check, 21, 24, 2
'    Loop until page_check = "THIS IS THE LAST PAGE"
'    END IF 
'    End If
'Next 





'For each row in the sheet
excel_row = 2 'start at row 2, row 1 is the header
ObjExcel.Sheets(2).Activate
Do while ObjExcel.Cells(excel_row, case_col).value <> ""
    MAXIS_case_number = ObjExcel.Cells(excel_row, case_col).value
    Call navigate_to_MAXIS_screen_review_priv("STAT", "AREP", is_this_priv)
    If is_this_priv = false Then
        EMREadScreen panel_exists, 1, 2, 78
        If panel_exists <> "0" Then
            EMReadscreen arep_name, 36, 4, 32
            arep_name = replace(arep_name, "_", "")
            arep_name = Trim(arep_name)
            EMReadscreen addr_line1, 21, 5, 32
            EMReadscreen addr_line2, 21, 6, 32
            address = addr_line1 & " " & addr_line2
            address = replace(address, "_", "")
            address = Trim(address)
            EMREadScreen phone, 21, 7, 32
            phone = replace(phone, "_", "")
            phone = replace(phone, "(", "")
            phone = replace(phone, ")", "")
            phone = replace(phone, " ", "")
            objexcel.Sheets(2).Cells(excel_row, arep_name_col) = arep_name
            objexcel.Sheets(2).Cells(excel_row, arep_addr_col) = address
            objexcel.Sheets(2).Cells(excel_row, arep_phone_col) = phone
        End If
    Else
        objExcel.Sheets(2).Cells(excel_row, arep_name_col) = "Priv Case"
    End If 
    excel_row = excel_row + 1
Loop
'If the case number is not blank, then
'   go to the case number
'   Read the AREP name / address
'   If the AREP name is not blank, then
'      Add to sheet
   ' End If 
'End If
   'Next 

script_end_procedure()
