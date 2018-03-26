'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - CLOSE GRH RATE 2 IN MMIS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 300                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
'END OF stats block==============================================================================================

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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/23/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

Function MMIS_panel_check(panel_name)
	Do 
		EMReadScreen panel_check, 4, 1, 51
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

function navigate_to_MMIS_region(group_security_selection)
'--- This function is to be used when navigating to MMIS from another function in BlueZone (MAXIS, PRISM, INFOPAC, etc.)
'~~~~~ group_security_selection: region of MMIS to access - programed options are "CTY ELIG STAFF/UPDATE", "GRH UPDATE", "GRH INQUIRY", "MMIS MCRE"
'===== Keywords: MMIS, navigate
	attn
	Do
		EMReadScreen MAI_check, 3, 1, 33
		If MAI_check <> "MAI" then EMWaitReady 1, 1
	Loop until MAI_check = "MAI"

	EMReadScreen mmis_check, 7, 15, 15
	IF mmis_check = "RUNNING" THEN
		EMWriteScreen "10", 2, 15
		transmit
	ELSE
		EMConnect"A"
		attn
		EMReadScreen mmis_check, 7, 15, 15
		IF mmis_check = "RUNNING" THEN
			EMWriteScreen "10", 2, 15
			transmit
		ELSE
			EMConnect"B"
			attn
			EMReadScreen mmis_b_check, 7, 15, 15
			IF mmis_b_check <> "RUNNING" THEN
				script_end_procedure("You do not appear to have MMIS running. This script will now stop. Please make sure you have an active version of MMIS and re-run the script.")
			ELSE
				EMWriteScreen "10", 2, 15
				transmit
			END IF
		END IF
	END IF

	DO
		PF6
		EMReadScreen password_prompt, 38, 2, 23
		IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then 
			Do 
				CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	 		Loop until are_we_passworded_out = false					'loops until user passwords back in
		End if 
		EMReadScreen session_start, 18, 1, 7
	LOOP UNTIL session_start = "SESSION TERMINATED"

	'Getting back in to MMIS and trasmitting past the warning screen (workers should already have accepted the warning when they logged themselves into MMIS the first time, yo.
	EMWriteScreen "MW00", 1, 2
	transmit
	transmit

	group_security_selection = UCASE(group_security_selection)

	EMReadScreen MMIS_menu, 24, 3, 30
	If MMIS_menu <> "GROUP SECURITY SELECTION" Then
		EMReadScreen mmis_group_selection, 4, 1, 65
		EMReadScreen mmis_group_type, 4, 1, 57

		correct_group = FALSE

		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			mmis_group_selection_part = left(mmis_group_selection, 2)

			If mmis_group_selection_part = "C3" Then correct_group = TRUE
			If mmis_group_selection_part = "C4" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the County Eligibility region. The script will now stop.")

		Case "GRH UPDATE"
			If mmis_group_selection  = "GRHU" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Update region. The script will now stop.")

		Case "GRH INQUIRY"
			If mmis_group_selection  = "GRHI" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the GRH Inquiry region. The script will now stop.")

		Case "MMIS MCRE"
			If mmis_group_selection  = "EK01" Then correct_group = TRUE
			If mmis_group_selection  = "EKIQ" Then correct_group = TRUE

			If correct_group = FALSE Then script_end_procedure("It does not appear you have access to the correct region of MMIS. This script requires access to the MCRE region. The script will now stop.")

		End Select

	Else
		Select Case group_security_selection

		Case "CTY ELIG STAFF/UPDATE"
			row = 1
			col = 1
			EMSearch " C3", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch " C4", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the County Eligibility area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH UPDATE"
			row = 1
			col = 1
			EMSearch "GRHU", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "GRH INQUIRY"
			row = 1
			col = 1
			EMSearch "GRHI", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				script_end_procedure("You do not appear to have access to the GRH Inquiry area of MMIS, this script requires access to this region. The script will now stop.")
			End If

			'Now it finds the pror authorization application feature and selects it.
			row = 1
			col = 1
			EMSearch "PRIOR AUTHORIZATION   ", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		Case "MMIS MCRE"
			row = 1
			col = 1
			EMSearch "EK01", row, col
			If row <> 0 Then
				EMWriteScreen "x", row, 4
				transmit
			Else
				row = 1
				col = 1
				EMSearch "EKIQ", row, col
				If row <> 0 Then
					EMWriteScreen "x", row, 4
					transmit
				Else
					script_end_procedure("You do not appear to have access to the MCRE area of MMIS, this script requires access to this region. The script will now stop.")
				End If
			End If

			'Now it finds the recipient file application feature and selects it.
			row = 1
			col = 1
			EMSearch "RECIPIENT FILE APPLICATION", row, col
			EMWriteScreen "x", row, col - 3
			transmit

		End Select
	End If
end function

'----------------------------------------------------------------------------------------------------DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog info_dialog, 0, 0, 266, 115, "Close MMIS service agreements in MMIS"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a GRH only list is provided from REPT/EOMC at the end of a month. These are cases that need to close in MMIS."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
get_county_code

MAXIS_footer_month = CM_mo	'establishing footer month/year 
MAXIS_footer_year = CM_yr 

'Determing the last day of the month to use as the closure date in MMIS.
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_of_month = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string
last_date = datePart("D", last_day_of_month)
end_date = CM_mo & last_date & CM_yr 
'MsgBox last_day_of_month & vbcr & end_date

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\GRH\GRH EOMC 03-18.xlsx"
excel_row_to_test = 2

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

objExcel.Cells(1, 6).Value = "PMI"
objExcel.Cells(1, 7).Value = "Case status"

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT
 
DIM Update_MMIS_array()
ReDim Update_MMIS_array(5, 0)

'constants for array
const case_number	= 0
const clt_PMI 	    = 1
const rate_two 	    = 2
const update_MMIS 	= 3
const case_status 	= 4

'Now the script adds all the clients on the excel list into an array
'excel_row = 2 're-establishing the row to start checking the members for
excel_row = excel_row_to_test
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
	If MAXIS_case_number = "" then exit do
	'Adding client information to the array'
	ReDim Preserve Update_MMIS_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	Update_MMIS_array(case_number,	entry_record) = MAXIS_case_number	'The client information is added to the array'
	Update_MMIS_array(clt_PMI, 	entry_record) = total_units				'STATIC for now. TODO: remove static coding for action script 
	Update_MMIS_array(rate_two, 	entry_record) = False               'default to False 
	Update_MMIS_array(update_MMIS, 	entry_record) = False				'This is the default, this may be changed as info is checked'
	Update_MMIS_array(case_status, 	entry_record) = ""					'This is the default, this may be changed as info is checked'
	
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop
'msgbox entry_record
'
back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

For item = 0 to UBound(Update_MMIS_array, 2)
	MAXIS_case_number = Update_MMIS_array(case_number ,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
		'msgbox "PRIV case, cannot access/update."
		Update_MMIS_array(rate_two, item) = False  	
		Update_MMIS_array(case_status, item) = "PRIV case, cannot access/update." 
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else 
		EMReadScreen grh_status, 4, 9, 74		'Ensuring that the case is active on GRH. If not, case will not be updated in MMIS. 
		If grh_status <> "ACTV" then 
			'msgbox "GRH case status is " & grh_status
			Update_MMIS_array(rate_two, item) = False 
			Update_MMIS_array(case_status, item) = "GRH case status is " & grh_status 	  
		Else 
			Update_MMIS_array(rate_two, item) = True  	
		End if
	End if 
    
    EMReadscreen current_county, 4, 21, 21
    If lcase(current_county) <> worker_county_code then 
        Update_MMIS_array(rate_two, item) = False 
        Update_MMIS_array(case_status, item) = "Out-of-county case."
    End if  
    
	Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there. 
	
	'----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the FACI dates match.
	If Update_MMIS_array(rate_two, item) = True then 
        Call navigate_to_MAXIS_screen ("STAT", "SSRT")				
        call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
        
        EMReadScreen SSRT_total_check, 1, 2, 78
        If SSRT_total_check = "0" then
            'msgbox "SSRT panel needs to be created."
            Update_MMIS_array(rate_two, item) = False  
            Update_MMIS_array(case_status, item) = "Case is not Rate 2."
        Else
            Update_MMIS_array(rate_two, item) = True 
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            EMReadScreen client_PMI, 8, 4, 46
            client_PMI = trim(client_PMI)
            client_PMI = right("00000000" & client_pmi, 8)
            Update_MMIS_array(clt_PMI, item) = client_pmi
            'msgbox MAXIS_case_number & vbcr & client_PMI
        End if 
    End if 
    
    If Update_MMIS_array(rate_two, item) = True then
		'----------------------------------------------------------------------------------------------------DISA: ensuring that client is not on a waiver. If they are, they should not be rate 2.
        Call navigate_to_MAXIS_screen("STAT", "DISA")
		Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01. 
		EMReadScreen waiver_type, 1, 14, 59
		If waiver_type <> "_" then 
			'msgbox "Client is active on a waiver. Should not be Rate 2."
			Update_MMIS_array(case_status, item) = "Client is active on a waiver. Should not be Rate 2."
			Update_MMIS_array(rate_two, item) = False 
		End if
	End if 	
Next 	

msgbox "Going into MMIS. Press OK to continue."

'----------------------------------------------------------------------------------------------------MMIS portion of the script
For item = 0 to UBound(Update_MMIS_array, 2)
	MAXIS_case_number = Update_MMIS_array(case_number,	item) 
	client_PMI = 		Update_MMIS_array(clt_PMI, 		item) 
	'msgbox MAXIS_case_number
	If Update_MMIS_array(rate_two, item) = True then
		'msgbox "True"
		Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior autorization area
		Call MMIS_panel_check("AKEY")				'ensuring we are on the right MMIS screen
	    EmWriteScreen client_PMI, 10, 36
	    Call write_value_and_transmit("C", 3, 22)	'Checking to make sure that more than one agreement is not listed by trying to change (C) the information for the PMI selected.
		EMReadScreen active_agreement, 12, 24, 2
		'msgbox active_agreement & vbcr & MAXIS_case_number
	    If active_agreement = "NO DOCUMENTS" then 
            Update_MMIS_array(update_MMIS, item) = False	
            Update_MMIS_array(case_status, item) = "Agreement for this PMI not found in MMIS."
        Else    
			EMReadScreen AGMT_status, 31, 3, 19 
			AGMT_status = trim(AGMT_status)
			If AGMT_status = "START DT:        END DT:" then 
			    Update_MMIS_array(update_MMIS, item) = False	
	    	    Update_MMIS_array(case_status, item) = "More than one service agreement exists in MMIS. Update manually."
			    PF3
            Else 
			    '----------------------------------------------------------------------------------------------------ASA1 screen
			    Call MMIS_panel_check("ASA1")				'ensuring we are on the right MMIS screen
		        EMReadScreen ASA1_end_date, 6, 4, 71
                If ASA1_end_date <> end_date then
                    Call write_value_and_transmit(end_date, 4, 71)				'End date is static for the BULK conversion. TODO: change to date_out which will match the FACI dates. 
		        ELSE
                    Transmit
                End if 
            
			    Call MMIS_panel_check("ASA2")				'ensuring we are on the right MMIS screen
			    transmit 	'no action required on ASA2
			    '----------------------------------------------------------------------------------------------------ASA3 screen
			    Call MMIS_panel_check("ASA3")				'ensuring we are on the right MMIS screen	
                EMReadScreen ASA3_end_date, 6, 8, 67
                If ASA3_end_date <> end_date then 
                    EMWriteScreen end_date, 8, 67
                    EMReadScreen start_month, 2, 8, 60
                    EMReadScreen start_day , 2, 8, 62
                    EMReadScreen start_year , 2, 8, 64
                    start_date = start_month & "/" & start_day & "/" & start_year
                    total_units = datediff("d", start_date, last_day_of_month) + 1
                    'msgbox total_units
                    
                    Call clear_line_of_text(9, 60)
                    EmWriteScreen total_units, 9, 60
			        PF3 '	to save changes 
                    Call MMIS_panel_check("AKEY")		'ensuring we are on the right MMIS screen
    			             
    			    EMReadscreen approval_message, 16, 24, 2
    			    If approval_message = "ACTION COMPLETED" then 
                        Update_MMIS_array(update_MMIS, item) = True 
                        Update_MMIS_array(case_status, item) = "SSR end date in MMIS updated to " & last_day_of_month	
                    Else
                        PF6
                        Update_MMIS_array(update_MMIS, item) = False 
                        Update_MMIS_array(case_status, item) = "Check case in MMIS. May not have updated, review manually."
    			    End if
			    Else 
                    Update_MMIS_array(update_MMIS, item) = False 
                    Update_MMIS_array(case_status, item) = "MMIS already updated for closure."	
                End if     
            End if 
		End if 		
	End if     
Next

'----------------------------------------------------------------------------------------------------EXCEL export
excel_row = excel_row_to_test

'Export informaiton to Excel re: case status
For item = 0 to UBound(Update_MMIS_array, 2)
	objExcel.Cells(excel_row, 6).Value = Update_MMIS_array(clt_PMI, item)
	objExcel.Cells(excel_row, 7).Value = Update_MMIS_array(case_status, item)
	excel_row = excel_row + 1
Next 

'formatting the cells
FOR i = 1 to 7
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

''----------------------------------------------------------------------------------------------------MAXIS 
Do 
	Call navigate_to_MAXIS("")
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("CASE", "NOTE")
'msgbox "Are we in MAXIS again?"
'----------------------------------------------------------------------------------------------------CASE NOTE
'Make the script case note
For item = 0 to UBound(Update_MMIS_array, 2)
	If Update_MMIS_array(update_MMIS, 	item) = True then 
		MAXIS_case_number = 	Update_MMIS_array(case_number, 	item)
		Call start_a_blank_CASE_NOTE
		Call write_variable_in_CASE_NOTE("GRH Rate 2 SSR closed in MMIS eff " & last_day_of_month)
		Call write_variable_in_CASE_NOTE("* Case set to close on REPT/EOMC, but still active in MMIS.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE("Actions performed by BZ script, run by I. Ferris, QI team")
		PF3
	End if 
	'msgbox "Did case note save for " & MAXIS_CASE_NUMBER
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")