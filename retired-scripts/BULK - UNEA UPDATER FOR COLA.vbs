'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - UNEA UPDATER FOR COLA.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
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
call changelog_update("12/16/2019", "Updated for 2020 COLA to work off of DAIL. BOBI does not provide data elements.", "Ilse Ferris, Hennepin County")
call changelog_update("12/02/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'------------------------------------------------------------------------------------------------------establishing date variables
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

first_of_month = CM_mo & "/" & "1" & "/" & CM_yr 

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\BZ ongoing projects\COLA\COLA Increase Automation\UNEA BOBI Pull.xlsx"


'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 50, "Select the COLA income source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog
'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog Dialog1 
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog
do 
    dialog Dialog1 
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


objExcel.Cells(1, 10).Value = "2020 Amount"
objExcel.Cells(1, 11).Value = "COLA Disregard"
objExcel.Cells(1, 12).Value = "Last Update Date"
objExcel.Cells(1, 13).Value = "Case Status"
objExcel.Cells(1, 14).Value = "Case Notes"

FOR i = 1 to 14	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets up the array to store all the information for each client'
Dim UNEA_array()
ReDim UNEA_array (8, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_number   = 0			'Each of the case numbers will be stored at this position'
Const memb_num    	= 1
Const inc_type		= 2
Const update_num   	= 3
Const unea_amt 	  	= 4
Const cola_amt    	= 5
Const cola_dis      = 6
Const case_status  	= 7
Const case_notes   	= 8

'Now the script adds all the clients on the excel list into an array
excel_row = excel_row_to_start 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
	    
	member_number = objExcel.cells(excel_row, 7).value	'establishes client SSN
    member_number = "0" & right(member_number, 2)

	income_type = objExcel.cells(excel_row, 8).value	
    income_type = trim(income_type)
    
    current_unea = objExcel.cells(excel_row, 9).value	
    current_unea = trim(current_unea)
    
	'Adding client information to the array'
	ReDim Preserve UNEA_array(8, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	UNEA_array (case_number,entry_record) = MAXIS_case_number		'The client information is added to the array'
	UNEA_array (memb_num,  	entry_record) = member_number
	UNEA_array (inc_type, 	entry_record) = income_type
    UNEA_array (update_num,	entry_record) = ""
	UNEA_array (unea_amt,   entry_record) = current_unea
	UNEA_array (cola_amt,   entry_record) = ""
    UNEA_array (cola_dis,   entry_record) = ""
	UNEA_array (case_status, entry_record) = ""
	UNEA_array (case_notes,  entry_record) = ""
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	Stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop

'msgbox entry_record

excel_row = excel_row_to_start
back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43		'Writes in Current month plus one
EMWriteScreen MAXIS_footer_year, 20, 46		'Writes in Current month plus one's year

For i = 0 to Ubound(UNEA_array, 2)
	'Establishing values for each case in the array of cases 
	MAXIS_case_number	= UNEA_array (case_number, i)
	income_type 		= UNEA_array (inc_type, i)
    member_number       = UNEA_array (memb_num, i)
	
    'msgbox MAXIS_case_number & vbcr & income_type & vbcr & member_number
    
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
    If PRIV_check = "PRIV" then
        UNEA_array(case_status, i) = "Error"
        UNEA_array(case_notes, i) = "Case is privileged."
        income_panel_found = false 
        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
        Do
            back_to_self
            EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
            If SELF_screen_check <> "SELF" then PF3
        LOOP until SELF_screen_check = "SELF"
        Call clear_line_of_text(18, 43)		'clears the case number
        transmit
    elseif income_type = "38" then 
        UNEA_array(case_status, i) = "Error"
        UNEA_array(case_notes, i) = "No COLA for VA A&A."	'Explanation for the rejected report'              
        income_panel_found = false 
    Else 
        MAXIS_background_check()
            
        EMReadScreen active_case, 8, 8, 9
        If active_case = "INACTIVE" then 
            UNEA_array(case_status, i) = "Error"
            UNEA_array(case_notes, i) = "Case is inactive."
            income_panel_found = false 
        Else 
	        'Checking the SNAP status 
	        Call navigate_to_MAXIS_screen("STAT", "PROG")		
	        EMReadscreen county_code, 2, 21, 23
	        If county_code <> "27" then 
	         	UNEA_array(case_status, i) = "Error"
	         	UNEA_array(case_notes, i) = "Not Hennepin County case, county code is: " & county_code	'Explanation for the rejected report'
	      	    income_panel_found = false 
	        Else 
	       		'Reads to see if the client is on SNAP 
	           	EMReadscreen SNAP_active, 4, 10, 74
	           	If SNAP_active = "ACTV" or SNAP_active = "REIN" then 
	       			update_SNAP = True
	       		Else 
	       			update_SNAP = false
	       		End if
	       			
	       		'Reads to see if the client is on HC
	       		EMReadScreen HC_active, 4, 12, 74
	       		If HC_active = "ACTV" or HC_active = "REIN" then 
	       			update_HC = True
	       		Else 
	       			update_HC = false
	       		End if 
                
	       		'handling for cases that do not have a completed HCRE panel
	       		PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	       		Do
	       			EMReadscreen HCRE_panel_check, 4, 2, 50
	       			If HCRE_panel_check = "HCRE" then
	       				PF10	'exists edit mode in cases where HCRE isn't complete for a member
	       				PF3
	       			END IF
	       		Loop until HCRE_panel_check <> "HCRE"
	       	End if 
        End if 
    End if 
    
    '----------------------------------------------------------------------------------------------------STAT UNEA PORTION
    If UNEA_array(case_status, i) = "" then 	
	    Call navigate_to_MAXIS_screen("STAT", "UNEA")
	    EMWriteScreen member_number, 20, 76
	    EMWriteScreen "01", 20, 79				'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	    transmit
	     
        'msgbox member_number 
	    EMReadScreen total_amt_of_panels, 1, 2, 78	'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
	    If total_amt_of_panels = "0" then 
	     	UNEA_array(case_status, i) = "Error"
	     	UNEA_array(case_notes, i) = "UNEA panel not known. Review case, and update manually if applicable."	'Explanation for the rejected report'              
	     	income_panel_found = false 
	    Else 	
	     	Do
	     		EMReadScreen current_panel_number, 1, 2, 73
	     		EMReadScreen panel_income_type, 2, 5, 37
                'msgbox income_type & vbcr & panel_income_type
	     		If income_type = panel_income_type then
	                income_panel_found = true 
                    
                    EMReadScreen update_date, 8, 21, 55
                    update_date = replace(update_date, " ", "/")
                    UNEA_array(update_num,  i) = update_date
                    If update_date = current_date then 
                        'msgbox "Case updated today: " & update_date
                        UNEA_array(case_status, i) = "Error"
                        UNEA_array(case_notes, i) = "Case updated today."	'Explanation for the rejected report'              
                        income_panel_found = false
                    Elseif cdate(update_date) > cdate(first_of_month) then 
                        'msgbox "updated this month"
                        UNEA_array(case_status, i) = "Error"
                        UNEA_array(case_notes, i) = "Case updated in 12/2019"	'Explanation for the rejected report'              
                        income_panel_found = false
                    Else     
                        EMReadscreen prospective_amt, 8, 13, 68
                        prospective_amt = trim(replace(prospective_amt, "_", ""))
                        'msgbox prospective_amt
                        
                        'not processing $90 VA amounts as these may be VA A & A as well
                        IF prospective_amt = "90.00" THEN
                            'msgbox "amount is 90. Not processing."
                            UNEA_array(case_status, i) = "Error"
                            UNEA_array(case_notes, i) = "VA income is $90. Check manually for A & A."	'Explanation for the rejected report'              
                            income_panel_found = false 
                        elseif prospective_amt = "0.00" THEN
                            'msgbox "amount is 0. Not processing."
                            UNEA_array(case_status, i) = "Error"
                            UNEA_array(case_notes, i) = "Income is 0. Review case."	'Explanation for the rejected report'              
                            income_panel_found = false 
                        Else 
                            PF9
                            If income_type = "11" then cola_muliplier = .016     'VA (1.6%)
                            If income_type = "12" then cola_muliplier = .016     'VA (1.6%)
                            If income_type = "13" then cola_muliplier = .016     'VA (1.6%)
                            If income_type = "38" then cola_muliplier = 0       'VA A & A - no updates for 2019
                            If income_type = "16" then cola_muliplier = .016     'Railroad (1.6%)
                            'If income_type = "17" then cola_muliplier = .0       'General PERA (1.4%)
                            
                            'Figuring out the calculations
                            increase_amt = prospective_amt * cola_muliplier
                            increase_amt = round(increase_amt, 2)
                            cola_amount = prospective_amt + increase_amt
                            cola_amount = round(cola_amount, 2)
                            
                            'MsgBox prospective_amt & vbcr & cola_amount & vbcr & increase_amt
                            
                            UNEA_array(unea_amt, i) = prospective_amt
                            UNEA_array(cola_amt, i) = cola_amount
                            UNEA_array(cola_dis, i) = increase_amt
                            
	                        'updates the SNAP PIC	
	                        If update_SNAP = true then 							
	                        	Call write_value_and_transmit("x", 10, 26)
	                        	Call create_MAXIS_friendly_date(date, 0, 5, 34)
	                        	EMWriteScreen "1", 5, 64							'code for pay frequency
	                        	row = 9											'blanking out the income fields on the PIC (just in case their is income listed there)
	                        	Do 
	                        		EMWriteScreen "__", row, 13
	                        		EMWriteScreen "__", row, 16
	                        		EMWriteScreen "__", row, 19
                                    EMWriteScreen "________", row, 25
	                                row = row + 1
	                            Loop until row = 14
	                                    	
	                            EMWriteScreen "________", 8, 66
                                EMWriteScreen cola_amount, 8, 66
	                                 
	                            Do 
	                            	transmit
	                            	EMReadscreen UNEA_panel, 4, 2, 48
	                            Loop until UNEA_panel = "UNEA"
	                        End if 	
	                        		
	                        'updates the HC pop up
	                        IF update_HC = true then							
	                        	Call write_value_and_transmit("x", 6, 56)
	                        	EMWriteScreen "________", 9, 65
	                        	EMWriteScreen cola_amount, 9, 65
	                        	EMWriteScreen "1", 10, 63							'code for pay frequency
	                        	Do 
	                        		transmit
	                        		EMReadscreen HC_popup, 9, 7, 41
	                        		If HC_popup = "HC Income" then transmit
	                        	Loop until HC_popup <> "HC Income"
	                        End if 
	     		                
	                        '----------------------------------------------------------------------------------------------------RETROSPECTIVE  					
	                        row = 13			'blanking out all retrospective UNEA fields
	                        DO 	
	                        	EMWriteScreen "__", row, 25
	                        	EMWriteScreen "__", row, 28
	                        	EMWriteScreen "__", row, 31
	                        	EMWriteScreen "________", row, 39
	                        	row = row + 1
	                        Loop until row = 18
	                        
	                        EMWriteScreen CM_minus_1_mo, 13, 25		'Entering the CM + 1 date 
	                        EMWriteScreen "01", 13, 28
	                        EMWriteScreen CM_minus_1_yr, 13, 31
	                        EMWriteScreen prospective_amt, 13, 39
	                        
	                        '----------------------------------------------------------------------------------------------------PROSPECTIVE 
	                        row = 13			'blanking out all prospective UNEA fields
	                        DO 
	                        	EMWriteScreen "__", row, 54
	                        	EMWriteScreen "__", row, 57
	                        	EMWriteScreen "__", row, 60
	                        	EMWriteScreen "________", row, 68
	                        	row = row + 1
	                        Loop until row = 18
	                    
	                        EMWriteScreen CM_plus_1_mo, 13, 54		'Entering the CM + 1 date 
	                        EMWriteScreen "01", 13, 57
	                        EMWriteScreen CM_plus_1_yr, 13, 60
	                        
	                        EMWriteScreen cola_amount, 13, 68		'Entering the income on the UNEA panel	
                            'msgbox cola_amount
                            '----------------------------------------------------------------------------------------------------UNEA panel updates
                            EMWriteScreen "6", 5, 65				'Verification code for 'worker initiated verification'
                            EMWriteScreen "________", 10, 67
                            'EmWriteScreen increase_amt, 10, 67      'Entering the cola disregard
                            transmit
                            PF3 'to exit the UNEA panel
	                        income_panel_found = True
	                        exit do 
                        End if 
                    End if 
                End if 
                transmit	'looking for another UNEA panel
                income_panel_found = false	
	     	Loop until current_panel_number = total_amt_of_panels
	     	
	     	'If income_panel_found <> true then 
	     	'	UNEA_array(case_status, i) = "Error"
	     	'	UNEA_array(case_notes, i) = "Unable to find UNEA panel."	'Explanation for the rejected report'
	     	'End if 
	        back_to_self		'to clear WRAP panel
	    End if 
	End if 
    
	IF income_panel_found = true then		
        If income_type = "11" then note_header = "VA Disa"
        If income_type = "12" then note_header = "VA Pension"
        If income_type = "13" then note_header = "VA Other"
        If income_type = "16" then note_header = "Railroad"
        
        increase_perc = "1.6%"
         
	    '----------------------------------------------------------------------------------------------------THE CASE NOTE
	    Call navigate_to_MAXIS_screen("CASE", "NOTE")
        'msgbox "Ready to case note"
        PF9
	    Call write_variable_in_CASE_NOTE("*" & note_header & " income update for 2020 COLA*")
	    Call write_variable_in_CASE_NOTE("COLA increased by $" & increase_amt & "(" & increase_perc & ")")
	    Call write_variable_in_CASE_NOTE("New UNEA amount: $" & cola_amount)
        Call write_variable_in_case_note("UNEA panel updated for this income type.")
		Call write_variable_in_case_note("---")
		Call write_variable_in_case_note(worker_signature)
		
		'ensuring that the case note saved. If not, adding it to the notes for the user to review. 
		PF3
		EMReadScreen note_date, 8, 5, 6
		If note_date <> current_date then 
			UNEA_array(case_status, i) = "Error"
			UNEA_array(case_notes, i) = "Case note does not appear to have been saved."	'Explanation for the rejected report'
	    Else 
            UNEA_array(case_status, i) = "Case updated"
            UNEA_array(case_notes, i) = ""	'Explanation for the rejected report' 
		End if 	
	End if
    
    
    ObjExcel.Cells(Excel_row, 10).Value = UNEA_array(cola_amt,    i)
    ObjExcel.Cells(Excel_row, 11).Value = UNEA_array(cola_dis,    i)
    ObjExcel.Cells(Excel_row, 12).Value = UNEA_array(update_num,  i)
    ObjExcel.Cells(Excel_row, 13).Value = UNEA_array(case_status, i)
    ObjExcel.Cells(Excel_row, 14).Value = UNEA_array(case_notes,  i)
    Excel_row = Excel_row + 1   
    
    prospective_amt = ""
    cola_amount = ""
    cola_muliplier = ""
    increase_amt = ""
Next    

FOR i = 1 to 14	'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Stats_counter = stats_counter + 1
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")