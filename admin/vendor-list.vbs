'STATS GATHERING----------------------------------------------------------------------------------------------------
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 30         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
BeginDialog vendor_dialog, 0, 0, 296, 105, "Vendor list"
  DropListBox 100, 75, 75, 15, "Select one..."+chr(9)+"GRH vendors"+chr(9)+"GRH vendor info"+chr(9)+"Non-GRH vendors", option_list
  ButtonGroup ButtonPressed
    OkButton 185, 75, 50, 15
    CancelButton 240, 75, 50, 15
  GroupBox 10, 10, 280, 55, "About this script:"
  Text 15, 25, 265, 20, "This script will gather the vendors for a specific county. Duplicates may appear. You will want to remove them from Excel."
  Text 35, 50, 225, 10, " Please shut down your VGO (not pause it), and press OK to continue."
  Text 20, 80, 75, 10, "Select a vendor option:"
EndDialog

'The script----------------------------------------------------------------------------------------------------
EMConnect ""		'Connecting to BlueZone

'The main dialog
Do
	Do
		dialog vendor_dialog
        If ButtonPressed = 0 then StopScript										
	LOOP until ButtonPressed = -1					'This is the OK button
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

If option_list <> "GRH vendor info" then 
'Adding script inforamtional data AND saving and closing actions----------------------------------------------------------------------------------------------------
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True
    
    'Setting the Excel rows with variables
    ObjExcel.Cells(1, 1).Value = "Vendor #"
    ObjExcel.Cells(1, 2).Value = "Vendor Name"
    ObjExcel.Cells(1, 3).Value = "Address"
    ObjExcel.Cells(1, 4).Value = "Status"
    
    FOR i = 1 to 4	'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
    excel_row = 2
    
    Call check_for_MAXIS(False)
    Call navigate_to_MAXIS_screen("MONY", "VNDS")
    EmWriteScreen "A", 5, 21 'only selects active vendors 
    If GRH_check = 1 then
        EmWriteScreen "Y", 5, 33
    ELSE
        EmWriteScreen "N", 5, 33
    End if 
    Call write_value_and_transmit("27", 5, 10) 'transmits to get vendors
    
    row = 8
    Do 
        EMReadScreen vendor_number, 8, row, 5
        EMReadScreen vendor_name, 30, row, 14
        EMReadScreen vendor_addr, 28, row, 45
        EMReadScreen vendor_status, 1, row, 79
    
        ObjExcel.Cells(excel_row, 1).Value = trim(vendor_number)
        ObjExcel.Cells(excel_row, 2).Value = trim(vendor_name)
        ObjExcel.Cells(excel_row, 3).Value = trim(vendor_addr)
        ObjExcel.Cells(excel_row, 4).Value = trim(vendor_status)
        excel_row = excel_row + 1
        row = row + 1
        If row = 18 then 
            PF8
            row = 8
        End if 
        EMReadScreen panel_limit, 12, 24, 2
        If panel_limit = "YOU CAN ONLY" then
            EMReadScreen last_vendor_name, 30, 17, 14
            If trim(last_vendor_name) = "MPLS PUBLIC HOUSING AUTH" then last_vendor_name = "MPLS PUBM"
            Call clear_line_of_text(4, 15)
            Call write_value_and_transmit(last_vendor_name, 4, 15)
            row = 9
        End if
    Loop until trim(vendor_number) = ""
    
    FOR i = 1 to 4										'formatting the cells'
    	objExcel.Columns(i).AutoFit()						'sizing the columns'
    NEXT
    script_end_procedure("All vendors have been added. Please clean up duplicate cases in Excel.")
End if 

If option_list = "GRH vendor info" then 
    'dialog
    BeginDialog file_select_dialog, 0, 0, 226, 50, "Select the GRH vendor file."
        ButtonGroup ButtonPressed
        PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
        OkButton 110, 30, 50, 15
        CancelButton 165, 30, 50, 15
        EditBox 5, 10, 165, 15, file_selection_path
    EndDialog

    file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Completed projects\Misc. completed projects\GRH vendor list 03-2018.xlsx"
    
    'dialog and dialog DO...Loop	
    Do
        'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
        'Show initial dialog
        Do
        	Dialog file_select_dialog
        	If ButtonPressed = cancel then stopscript
        	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        Loop until ButtonPressed = OK and file_selection_path <> ""
        If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
    
    objExcel.Cells(1, 5).Value = "NPI"
    objExcel.Cells(1, 6).Value = "Rate 2 SSR"
    
    FOR i = 1 to 6		'formatting the cells'
    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
    	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
    excel_row = 2
    Do 
        vendor_number = ObjExcel.Cells(excel_row, 1).Value
	    vendor_number = trim(vendor_number)
         
        '----------------------------------------------------------------------------------------------------VNDS/VND2
        Call Navigate_to_MAXIS_screen("MONY", "VNDS")
        Call write_value_and_transmit(vendor_number, 4, 59)
        Call write_value_and_transmit("VND2", 20, 70)
        EMReadScreen NPI_number, 10, 7, 41
        NPI_number = replace(NPI_number, "_", "")
        
        EmreadScreen SSR_check, 5, 16, 45
        EMReadScreen rate_two_check, 6, 15, 63
        
        If SSR_check = "(SSR)" then 
            EMReadScreen service_rate, 8, 16, 68		'Reading the service rate to input into Excel
        Elseif rate_two_check = "Rate 2" then
            EMReadScreen service_rate, 8, 15, 72		'Reading the service rate to input into Excel 
        Else 
            service_rate = ""
        End if 

        ObjExcel.Cells(excel_row, 5).Value = trim(NPI_number)	
        ObjExcel.Cells(excel_row, 6).Value = trim(service_rate)
        
        stats_counter = stats_counter + 1
        excel_row = excel_row + 1
    Loop until trim(vendor_number) = ""
    
    FOR i = 1 to 6										'formatting the cells'
        objExcel.Columns(i).AutoFit()						'sizing the columns'
    NEXT
    
    Stats_counter = stats_counter - 1
    script_end_procedure("All vendors have been updated. Please review list.")
End if 
