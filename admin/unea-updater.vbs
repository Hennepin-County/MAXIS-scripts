'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - UNEA UPDATER.vbs"
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
CALL changelog_update("11/01/2020", "Removed COLA information, it's not applicable, and updated data columns.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/16/2019", "Updated data columns based on current data pull.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/12/2019", "Updated text for case note re: veterans services.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/04/2019", "Updated column numbers. New information is being pulled into the report.", "Ilse Ferris, Hennepin County")
CALL changelog_update("06/08/2018", "Removed custom function. This is now in the HC Functions Library. Updated back end dialog functionality and updated ", "Ilse Ferris, Hennepin County")
CALL changelog_update("02/05/2018", "Added additional handling for SPEC/MEMO sending, data validation and comments.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC/MEMO. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("07/28/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'------------------------------------------------------------------------------------------------------establishing date variables
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 110, "UNEA Updater"
  GroupBox 10, 5, 230, 80, "Using this script:"
  Text 20, 20, 210, 20, "This script should be used when a list of UNEA income has been provided and verified through internal sources."
  DropListBox 120, 50, 115, 15, "Select one..."+chr(9)+"Unemployment (UC)"+chr(9)+"Veterans (VA)", type_selection
  ButtonGroup ButtonPressed
  PushButton 195, 65, 40, 15, "Browse...", select_a_file_button
  OkButton 150, 90, 40, 15
  CancelButton 195, 90, 40, 15
  EditBox 20, 65, 170, 15, file_selection_path
  Text 20, 50, 95, 10, "Select the processing option:"
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
        If type_selection = "Select one..." then err_msg = err_msg & vbNewLine & "Select the income type."
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data."
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Creating an array of letters to loop through
col_ind = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG~AH~AI~AJ~AK~AL~AM~AN~AO~AP~AQ~AR~AS~AT~AU~AV~AW~AX~AY~AZ"
col_array = split(col_ind, "~")
'setting the start of the list of column options
column_list = "Select One..."
cell_val = 1        'starting the value for reading the top cell of each column to use header information

'looping through the array
For each letter in col_array
    col_header = UCase(objExcel.Cells(1, cell_val).Value)
    col_header = trim(col_header)

    If col_header <> ""  then                                              'if the column is not blank - add to dropdown
        column_list = column_list & chr(9) & letter & " - " & col_header
        If col_header = "CASE NUMBER" or col_header = "CASENUMBER" then case_number_col = letter & " - " & col_header
        If col_header = "PERSONID" or col_header = "PMI" then pmi_col = letter & " - " & col_header     
        If col_header = "INCOME TYPE CODE" or col_header = "INCOMETYPECODE" then income_type_col = letter & " - " & col_header  
        If col_header = "CLAIM NBR" or col_header = "CLAIMNBR" then claim_col = letter & " - " & col_header  
        If instr(col_header) = "INCORRECT" then act_claim_col = letter & " - " & col_header  
        If col_header = "AMT" or col_header = "WEEKLY AMT" then unea_col = letter & " - " & col_header  
        If col_header = "ACCT BALANCE" then balance_col = letter & " - " & col_header    
        If col_header = "CASE STATUS" or col_header = "STATUS" then status_col = letter & " - " & col_header  
        If col_header = "NOTES" or col_header = "CASE NOTES" then notes_col = letter & " - " & col_header
    Else
        last_col = letter       'setting this for adding additional columns with information
        Exit For
    End If
    cell_val = cell_val + 1
Next

'Next dialog determines the column the case numbers are in and the type of notification to be sent.
'Defining the dialog here so that the list of columns can be dynamically generated
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 135, "Select Data Locations"
  DropListBox 160, 70, 100, 45, column_list, case_number_column
  ButtonGroup ButtonPressed
    OkButton 155, 115, 50, 15
    CancelButton 210, 115, 50, 15
  Text 10, 10, 245, 20, "Check the Excel File that has been opened. Be sure it is the correct file to run at this time."
  Text 10, 35, 245, 30, "Choose the column that has all the case numbers listed and select which type of notice should be sent. The script will run very differently based on these answers."
  Text 10, 70, 145, 10, "Indicate the column with the case numbers:"
  Text 10, 95, 140, 10, "Which type of notice do you want to send?"
  Text 10, 120, 60, 10, "Excel row to start:"
EndDialog

'Displaying the dialog to select the correct column and type of notice.
Do
    Dialog Dialog1
    If ButtonPressed = cancel then stopscript
Loop until case_number_column <> "Select One..." AND notice_type <> "Select One..."

call back_to_self
EMReadScreen mx_region, 10, 22, 48

If mx_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are attempting to have the script send notices for these cases." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
    If continue_in_inquiry = vbNo Then script_end_procedure("Live script run was attempted in Inquiry and aborted.")
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 196, 150, "Confirm Selections"
  Text 10, 10, 175, 20, "You are running a BULK script that will send notices. Review the Excel Spreadsheet that opened."
  Text 10, 35, 175, 10, "Worksheet selected: " & scenario_dropdown
  Text 10, 55, 175, 10, "Case Number Column: " & case_number_column
  Text 10, 75, 175, 10, "Notice to be sent: " & notice_type
  Text 10, 90, 180, 35, "This is a long running script and you will be unable to use any Excel document or the current session of MAXIS while the script runs. Review the selected options to be sure the script will takethe correct action."
  ButtonGroup ButtonPressed
    PushButton 80, 130, 50, 15, "Confirm", cnfrm_btn
    CancelButton 140, 130, 50, 15
EndDialog

Do
    Dialog Dialog1
    cancel_without_confirmation
Loop until buttonpressed = cnfrm_btn

'Setting the Excel Columns 
case_number_col     =  
pmi_col             =
income_type_col     = 
claim_col           = 
act_claim_col       = 
unea_col            = 
status_col          = 
notes_col           =
balance_col         = 

'Sets up the array to store all the information for each client'
Dim UNEA_array()
ReDim UNEA_array (9, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num    	= 0			'Each of the case numbers will be stored at this position'
Const clt_pmi     	= 1
Const inc_type		= 2
Const claim_num   	= 3
Const act_claim		= 4
Const unea_amt 	  	= 5
Const act_status  	= 6
Const act_notes   	= 7
Const send_memo     = 8
Const acct_bal      = 9

'converting the column letter to a number because cell values are called by number
col = left(case_number_column, 1)
call convert_excel_letter_to_excel_number(col)

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 1).Value          're-establishing the case numbers for functions to use
	If MAXIS_case_number = "" then exit do
	MAXIS_case_number = trim(MAXIS_case_number)

	client_PMI = objExcel.cells(excel_row, 2).value	'establishes client SSN
	'removing the 0's from the PMI number to match the formatting from MAXIS
	Do
		client_PMI = trim(client_PMI)
		If left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) - 1)
	Loop until left(client_PMI, 1) <> "0"

	income_type  	= objExcel.cells(excel_row, income_col).value 
	claim_number 	= objExcel.cells(excel_row, claim_col).value 
    actual_claim 	= objExcel.cells(excel_row, act_claim_col).value 
	unea_amount	 	= objExcel.cells(excel_row, unea_col).value 
    account_balance = objExcel.cells(excel_row, balance_col).value 

	'Adding client information to the array'
	ReDim Preserve UNEA_array(9, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	UNEA_array (case_num, 	entry_record) = trim(MAXIS_case_number)		'The client information is added to the array'
	UNEA_array (clt_PMI,  	entry_record) = trim(client_PMI)
	UNEA_array (inc_type, 	entry_record) = trim(income_type)
    UNEA_array (claim_num,	entry_record) = trim(claim_number)
	UNEA_array (act_claim, 	entry_record) = trim(actual_claim)
	UNEA_array (unea_amt,   entry_record) = trim(unea_amount)
	UNEA_array (act_status, entry_record) = ""
	UNEA_array (act_notes,  entry_record) = ""
    UNEA_array (send_memo,  entry_record) = False
    UNEA_array (acct_bal,   entry_record) = trim(account_balance)
	entry_record = entry_record + 1			'This increments to the next entry in the array'
	Stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop

back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43		'Writes in Current month plus one
EMWriteScreen MAXIS_footer_year, 20, 46		'Writes in Current month plus one's year

For i = 0 to Ubound(UNEA_array, 2)
	'Establishing values for each case in the array of cases
	MAXIS_case_number	= UNEA_array (case_num, i)
	client_PMI			= UNEA_array (clt_PMI, i)
	income_type 		= UNEA_array (inc_type, i)
    actual_claim        = UNEA_array (act_claim, i)
	unea_amount 		= UNEA_array (unea_amt, i)

	If unea_amount = "" or IsNumeric(unea_amount) = False then
		UNEA_array(act_status, i) = "Error"
		UNEA_array(act_notes, i) = "VA income amount is blank or is not numeric."
        UNEA_array(send_memo, i) = False
		income_panel_found = false
	Else
	    MAXIS_background_check()

        Call navigate_to_MAXIS_screen("CASE", "CURR")
        EMReadScreen active_case, 8, 8, 9
        If active_case = "INACTIVE" then
            UNEA_array(act_status, i) = "Error"
            UNEA_array(act_notes, i) = "Case is inactive."
            UNEA_array(send_memo, i) = False
            income_panel_found = false
        Else
	        'Checking the SNAP status
	        Call navigate_to_MAXIS_screen("STAT", "PROG")
	        EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	        If PRIV_check = "PRIV" then
	        	UNEA_array(act_status, i) = "Error"
	        	UNEA_array(act_notes, i) = "Case is privileged."
                UNEA_array(send_memo, i) = False
	        	income_panel_found = false

	        	'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
	        	Do
	        		back_to_self
	        		EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
	        		If SELF_screen_check <> "SELF" then PF3
	        	LOOP until SELF_screen_check = "SELF"
	        	EMWriteScreen "________", 18, 43		'clears the case number
	        	transmit
	        Else
	            EMReadscreen county_code, 2, 21, 23
	            If county_code <> "27" then
	            	UNEA_array(act_status, i) = "Error"
	            	UNEA_array(act_notes, i) = "Not Hennepin County case, county code is: " & county_code	'Explanation for the rejected report'
                    UNEA_array(send_memo, i) = False
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

	            	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	            	Do
	            		EMReadscreen client_PMI, 8, 4, 46
	            		client_PMI = trim(client_PMI)
	            		If client_PMI = UNEA_array(clt_PMI, i) then
	            			EMReadscreen member_number, 2, 4, 33
	        				exit do
	            		Else
	            			transmit
	            		END IF
	            		EMReadScreen MEMB_error, 5, 24, 2
	            	Loop until client_PMI = UNEA_array (clt_SSN, i) or MEMB_error = "ENTER"

	            	IF client_PMI <> UNEA_array(clt_PMI, i) then
	            		UNEA_array(act_status, i) = "Error"
	            		UNEA_array(act_notes, i) = "Unable to find person's member number."	'Explanation for the rejected report'
                        UNEA_array(send_memo, i) = False
	        			income_panel_found = false
	            	Else
	            		'STAT UNEA PORTION
	            		Call navigate_to_MAXIS_screen("STAT", "UNEA")
	        			EMWriteScreen member_number, 20, 76
	        			EMWriteScreen "01", 20, 79				'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	        			transmit

	        			EMReadScreen total_amt_of_panels, 1, 2, 78	'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
	        			If total_amt_of_panels = "0" then
	        				UNEA_array(act_status, i) = "Error"
	        				UNEA_array(act_notes, i) = "UNEA panel not known. Review case, and update manually if applicable."	'Explanation for the rejected report'
                            UNEA_array(send_memo, i) = False
	        				income_panel_found = false
	        			Else
	        				Do
	        					EMReadScreen current_panel_number, 1, 2, 73
	        					EMReadScreen income_type, 2, 5, 37
	        					If income_type = UNEA_array(inc_type, i) then
	        						income_panel_found = true
	        						PF9

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
	        							EMWriteScreen UNEA_array(unea_amt, i), 8, 66
	        							Do
	        								transmit
	        								EMReadscreen UNEA_panel, 4, 2, 48
	        							Loop until UNEA_panel = "UNEA"
	        						End if

	        						'updates the HC pop up
	        						IF update_HC = true then
	        							Call write_value_and_transmit("x", 6, 56)
	        							EMWriteScreen "________", 9, 65
	        							EMWriteScreen UNEA_array(unea_amt, i), 9, 65
	        							EMWriteScreen "1", 10, 63							'code for pay frequency
	        							Do
	        								transmit
	        								EMReadscreen HC_popup, 9, 7, 41
	        								If HC_popup = "HC Income" then transmit
	        							Loop until HC_popup <> "HC Income"
	        						End if
	        						'----------------------------------------------------------------------------------------------------UNEA panel updates
	        						EMWriteScreen "6", 5, 65				'Verification code for 'worker initiated verification'

	        						If UNEA_array(act_claim, i) <> "" then 		'If the case's claim number has been identified as being incorrect, the correct claim will be entered.
	        							EMWriteScreen "_______________", 6, 37
	        							EMWriteScreen UNEA_array(act_claim, i), 6, 37
	        						End if

	        						'----------------------------------------------------------------------------------------------------RETROSPECTIVE
	        						EMReadscreen prospective_amt, 8, 13, 68
	        						prospective_amt = replace(prospective_amt, "_", "")

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

	        						EMWriteScreen UNEA_array(unea_amt, i), 13, 68		'Entering the income on the UNEA panel
	        						transmit
	        						PF3 		'to exit the UNEA panel
	        						income_panel_found = True
	        						exit do
	        					Else
	        						transmit	'looking for another UNEA panel
	        					End if
	        				Loop until current_panel_number = total_amt_of_panels

	        				If income_panel_found <> true then
	        					UNEA_array(act_status, i) = "Error"
	        					UNEA_array(act_notes, i) = "Unable to find person's member number."	'Explanation for the rejected report'
	        					UNEA_array(send_memo, i) = False
	        				End if
	        				back_to_self		'to clear WRAP panel
	            		End if
	            	End if
	            End if
	        End if
        End if
	End if

	IF income_panel_found = true then
	    start_a_blank_CASE_NOTE
	    '----------------------------------------------------------------------------------------------------THE CASE NOTE
	    renewal_period = MAXIS_footer_month & "/" & MAXIS_footer_year		'establishing the renewal period for the header of the case note

	    start_a_blank_CASE_NOTE
	    Call write_variable_in_CASE_NOTE("*" & renewal_period & " recert accuracy update for VA income*")
	    Call write_variable_in_CASE_NOTE("Do not update the following info unless a new change has been reported.")
	    Call write_variable_in_CASE_NOTE("* VA income: $" & UNEA_array(unea_amt, i) & " monthly grant.")
		Call write_variable_in_CASE_NOTE("* VA income has been verified via phone by Hennepin County Veterans Service Office staff.")
		call write_variable_in_case_note("* SPEC/MEMO sent to promote Hennepin County Veterans Service Office.")

		call write_variable_in_case_note("---")
		call write_variable_in_case_note(worker_signature)

		'ensuring that the case note saved. If not, adding it to the notes for the user to review.
		PF3
		EMReadScreen note_date, 8, 5, 6
		If note_date <> current_date then
			UNEA_array(act_status, i) = "Error"
			UNEA_array(act_notes, i) = "Case note does not appear to have been saved."	'Explanation for the rejected report'
			UNEA_array(send_memo, i) = False
	    Else
            UNEA_array(act_status, i) = "Case updated"
            UNEA_array(act_notes, i) = ""	'Explanation for the rejected report'
			UNEA_array(send_memo, i) = True
		End if
	End if
Next

If type_selection = "Veterans (VA)" then 
    For i = 0 to Ubound(UNEA_array, 2)
        If UNEA_array(send_memo, i) = True then
            MAXIS_case_number = UNEA_array(case_num, i)
            Call MAXIS_background_check
            '----------------------------------------------------------------------------------------------------THE SPEC/MEMO
            Call start_a_new_spec_memo
            call navigate_to_MAXIS_screen("SPEC", "MEMO")		'Navigating to SPEC/MEMO
            'Writes the MEMO.
            call write_variable_in_SPEC_MEMO("If you have any questions about veterans benefits, please contact the Hennepin County Veterans Service Office at 612-348-3300. Veterans Services has staff at the Government Center, the South Minneapolis Human Service center, and Maple Grove. You may also make an appointment at a variety of regional locations.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("Even if you are already in receipt of compensation or pension, your benefit amount may be able to be increased.")
            Call write_variable_in_SPEC_MEMO("")
            Call write_variable_in_SPEC_MEMO("If you are interested in speaking with someone regarding Veterans benefits, or if you have questions about this notice, please call the Hennepin County Veterans Service Office at 612-348-3300. Thank you.")
            PF4			'Exits the MEMO
            EMReadScreen memo_sent, 8, 24, 2
            If memo_sent <> "NEW MEMO" then
                UNEA_array(act_status, i) = "Error"
                UNEA_array(act_notes, i) = "Does not appear that memo sent."	'Explanation for the rejected report'
                PF10
            End if
        End if
    Next
End if 

'Export data to Excel
excel_row = 2
For i = 0 to Ubound(UNEA_array, 2)
	ObjExcel.Cells(Excel_row,  9).Value = UNEA_array(act_status, i) '(Col I)
	ObjExcel.Cells(Excel_row, 10).Value = UNEA_array(act_notes,  i) '(Col J)
	Excel_row = Excel_row + 1
Next

Stats_counter = stats_counter + 1
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
