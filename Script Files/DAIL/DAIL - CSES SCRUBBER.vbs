'Required for statistical purposes=========================================================================================
name_of_script = "DAIL - CSES SCRUBBER.vbs"
start_time = timer
STATS_counter = 0              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.
STATS_manualtime = 54          'manual run time in seconds
STATS_denomination = "I"       'I is for each dail message
'END OF stats block========================================================================================================

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

'FUNCTIONS=================================================================================================================
FUNCTION create_mainframe_friendly_date(date_variable, screen_row, screen_col, year_type)
	var_month = datepart("m", date_variable)
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	EMWriteScreen var_month, screen_row, screen_col
	var_day = datepart("d", date_variable)
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	EMWriteScreen var_day, screen_row, screen_col + 3
	If year_type = "YY" then
		var_year = right(datepart("yyyy", date_variable), 2)
	ElseIf year_type = "YYYY" then
		var_year = datepart("yyyy", date_variable)
	Else
		MsgBox "Year type entered incorrectly. Fourth parameter of function create_mainframe_friendly_date should read ""YYYY"" or ""YY"". The script will now stop."
		StopScript
	END IF
	EMWriteScreen var_year, screen_row, screen_col + 6
END FUNCTION

'END FUNCTIONS=============================================================================================================

'DIALOGS===================================================================================================================
BeginDialog CSES_initial_dialog, 0, 0, 296, 40, "CSES Dialog"
  'CheckBox 5, 5, 290, 10, "Check here if you would like to see an Excel sheet of the CSES scrubber calculations.", excel_visible_checkbox
  EditBox 70, 20, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 20, 50, 15
    CancelButton 240, 20, 50, 15
  Text 5, 25, 65, 10, "Worker signature:"
EndDialog
'END DIALOGS===============================================================================================================

'VARIABLES AND CONSTANTS WE WANT TO USE====================================================================================
const excel_center = -4108		'This apparently means "centered" in Excel's VBA

'Message/disbursement columns (used by the Excel to store info)
const col_msg_number 				= 1
const col_PMI_number 				= 2
const col_HH_memb_number 			= 3
const col_amt_alloted 				= 4
const col_CS_type	 				= 5
const col_issue_date				= 6
const col_UNEA_panel				= 7
const col_open_program_titles		= 9
const col_open_program_status		= 10
const col_HH_memb_PMI_list_memb_num	= 12
const col_HH_memb_PMI_list_PMI		= 13

message_array = array()			'A blank array which will be used when we move info between sheets on Excel
'END VARIABLES=============================================================================================================

'CLASSES===================================================================================================================
'A MessageDetails class for use in sorting/filtering message details
class MessageDetails
	public MsgNum
	public PMINum
	public MEMBNum
	public AmtAlloted
	public CSType
	public IssueDate
	public UNEAPanel
end class
'END CLASSES===============================================================================================================

'THE SCRIPT================================================================================================================

'Connects to MAXIS
EMConnect ""

'If the worker signature is the Konami code (UUDDLRLRBA), developer mode will be triggered
If worker_signature = "UUDDLRLRBA" then
    MsgBox "Developer mode: ACTIVATED!"
    developer_mode = true           'This will be helpful later
    collecting_statistics = false   'Lets not collect this, shall we?		'<<<<CHECK THIS, I THINK THE VARIABLE IS CALLED SOMETHING DIFFERENT IN THE FUNCTION
End if

'Checks if you're on a TIKL, and asks if this is a training scenario
EMReadScreen TIKL_check, 4, 6, 6
If TIKL_check = "TIKL" then
    'TIKL_processing_confirmation = MsgBox("You seem to be running this on a TIKL. Are you testing the script? If you aren't, something has gone wrong.", vbYesNo)		<<<<RESET THIS PLEEEEEEEEEEEEEEEEEEEEEEEEEEEASE
    'If TIKL_processing_confirmation = vbNo then stopscript																												<<<<RESET THIS PLEEEEEEEEEEEEEEEEEEEEEEEEEEEASE
	message_type_code = "TIKL"		'Uses this later on to determine if we're on the right messages on a DAIL
Else
	message_type_code = "CSES"		'Uses this later on to determine if we're on the right messages on a DAIL
End if

'EXCEL BLOCK------------------------------
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = true
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = true
'END EXCEL BLOCK--------------------------

'Renames this sheet
ObjExcel.ActiveSheet.Name = "Message and Disb info"

'Headers for the Excel spreadsheet
ObjExcel.Range("A1:G1").Merge
ObjExcel.Cells(1, col_msg_number).Value 				= "Message/disbursement info"
ObjExcel.Cells(1, col_msg_number).Font.Bold 			= True
objExcel.Cells(1, col_msg_number).HorizontalAlignment 	= excel_center
ObjExcel.Cells(2, col_msg_number).Value 				= "Message #"
ObjExcel.Cells(2, col_PMI_number).Value 				= "PMI"
ObjExcel.Cells(2, col_HH_memb_number).Value				= "HH memb"
ObjExcel.Cells(2, col_amt_alloted).Value 				= "Amount alloted"
ObjExcel.Cells(2, col_CS_type).Value 					= "CS type"
ObjExcel.Cells(2, col_issue_date).Value 				= "Issue date"
ObjExcel.Cells(2, col_UNEA_panel).Value 				= "UNEA Panel"

'Make the headers bold
For i = 1 to 7
	objExcel.Cells(2, i).Font.Bold = TRUE
Next


'We need these variables for the next part
excel_row = 3 		'What row should Excel be on? Let's start with this one.
message_number = 1	'We want to count how many messages we process in here



'===================================================================================================================================READS EACH MESSAGE!
For MAXIS_row = 6 to 19			'<<<<<CHECK THIS AGAINST A FULL, ACTUAL FACTUAL DAIL
	EMReadScreen message_type_check, 4, MAXIS_row, 6				'Makes sure it's the right type of message
	If message_type_check <> message_type_code then exit for 		'This was determined above based on TIKL vs actual CSES messages. If we aren't on the right message, it will exit
	EMWriteScreen "x", MAXIS_row, 3									'Puts an 'X' on the DAIL message
	transmit														'Transmits

	'READS THE TYPE
	row = 1
	col = 1
	EMSearch "TYPE", row, col
	EMReadScreen CS_type, 2, row, col + 5

	'<<<<SPOUSAL HANDLING SHOULD GO HERE BUT FOR NOW I'M SKIPPING IT

	'REDECLARES THESE VARIABLES (TYPE IS IN A DIFFERENT PLACE THAN THE AMOUNT) '<<<CAN'T WE FLIP THIS AROUND MAYBE?
	row = 1
	col = 1

	'READS THE AMOUNT
	EMSearch "$", row, col
	EMReadScreen COEX_amt, 6, row, col + 1
	COEX_amt = Replace(COEX_amt, "F", "")		'I've seen an "F" in here and I'm not totally sure why

	'READS THE TOTAL NUMBER OF KIDS THIS RELATES TO (TO BE USED AS A DENOMINATOR IN OUR CALCULATIONS)
	EMSearch "CHILD(REN)", row, col
    EMReadScreen COEX_PMI_total, 1, row, col - 2

	'READS THE PMIS AND DATE FROM THE MESSAGE
	EMSearch " TO PMI(S): ", row, col										'First it finds the PMIs on the screen
	EMReadScreen issue_date, 				08, row, col - 8				'The date is always before the "TO PMI(S): " string, apparently
	EMReadScreen raw_PMI_numbers_initial, 	40, row, col + 12				'Reads the contents immediately after the "TO PMI(S): " string, because sometimes a PMI number sneaks in there
	EMReadScreen raw_PMI_numbers_overflow, 	70, row + 1, 5					'Reads the next line in its entirety (all that would be here are PMIs)
	raw_PMI_numbers = raw_PMI_numbers_initial & raw_PMI_numbers_overflow	'Concatenates the two strings together
	PMI_numbers_no_spaces = Replace(raw_PMI_numbers, " ", "")				'Removes spaces from the lines
	PMI_array = Split(PMI_numbers_no_spaces, ",")							'Splits PMIs into an array

	'ADDS THE INFO TO EXCEL BASED ON PMI
	For each PMI_number in PMI_array
    	ObjExcel.Cells(excel_row, col_msg_number).Value 	= message_number					'Each message is numbered in sequence
    	ObjExcel.Cells(excel_row, col_PMI_number).Value 	= PMI_number						'We want this PMI for obvious reasons
    	ObjExcel.Cells(excel_row, col_amt_alloted).Value 	= COEX_amt / COEX_PMI_total		'Amount / total recipients gives us the amount per recipient

		penny_issue_total_cell_amt_times_100 = (ObjExcel.Cells(excel_row, col_amt_alloted).Value) * 100 											'Grabs the amount to be evaluated multiplies by 100 to get rid of the first two digits of the decimal
		penny_issue_partial_pennies_from_cell = (penny_issue_total_cell_amt_times_100 - int(penny_issue_total_cell_amt_times_100) ) / 100			'Grabs the actual partial pennies by eliminating the integer from the previous value, then dividing by 100 to return it to the proper place in the decimal
		penny_issue_partial_pennies_total = penny_issue_partial_pennies_total + penny_issue_partial_pennies_from_cell 								'Adds the partial pennies to a new variable to be tacked on at the end
		ObjExcel.Cells(excel_row, col_amt_alloted).Value = ObjExcel.Cells(excel_row, col_amt_alloted).Value - penny_issue_partial_pennies_from_cell	'Updates the cell to eliminate the partial pennies


    	ObjExcel.Cells(excel_row, col_CS_type).Value 		= CS_type						'This is the type, and it's helpful to know this when we write to UNEA
    	ObjExcel.Cells(excel_row, col_issue_date).Value 	= issue_date						'The date it was issued
    	excel_row = excel_row + 1											'Increments up one in order to start on the next Excel row
    Next

	'Adding partial pennies to the member 01
	ObjExcel.Cells(3, col_amt_alloted).Value = ObjExcel.Cells(3, col_amt_alloted).Value + penny_issue_partial_pennies_total

	'Clearing this variable so we can start over again next message (next run through the loop)
	penny_issue_partial_pennies_total = 0

	'GETS OUT OF THE MESSAGE
	transmit

	'ADDS ONE TO THE MESSAGE NUMBER SO WE CAN KEEP A GOOD COUNT
	message_number = message_number + 1
Next




'===================================================================================================================================DETERMINING WHAT PROGRAMS ARE OPEN
'Navigates to CASE/CURR directly (the DAIL doesn't easily go back to the case-in-question when we use the custom function)
EMWriteScreen "h", 6, 3
transmit

'First, checks for inactive cases and just shuts down if it finds one
row = 1
col = 1
EMSearch "Case: INACTIVE", row, col
If row <> 0 then script_end_procedure("This case is inactive in MAXIS. The script will now stop.")

'Then it checks for HC active. For right now, it'll just create a pop-up saying it doesn't do anything for HC at present
row = 1
col = 1
EMSearch "HC:", row, col
'If row <> 0 then MsgBox "As of March 2016 the health care sections have been removed from the CSES Scrubber. Evaluate any health care ramifications manually."		'<<<<<<<RESET THIS PLEASE, OR PUT IT IN THE DIALOG
If row <> 0 then
	HC_active = True
Else
	HC_active = False
End if

'Then it checks for MFIP status
row = 1
col = 1
EMSearch "MFIP:", row, col
If row <> 0 then
	MFIP_active = True
Else
	MFIP_active = False
End if

'Then it checks for SNAP status
row = 1
col = 1
EMSearch "FS:", row, col
If row <> 0 then
	SNAP_active = True
Else
	SNAP_active = False
End if

'Writes program status to the Excel sheet, because it's prettier that way (and will be helpful for debugging)
ObjExcel.Cells(1, col_open_program_titles).Value 				= "CASE/CURR status"
ObjExcel.Cells(1, col_open_program_titles).Font.Bold 			= TRUE
objExcel.Cells(1, col_open_program_titles).HorizontalAlignment 	= excel_center
ObjExcel.Range("I1:J1").Merge
ObjExcel.Cells(2, col_open_program_titles).Value 				= "MFIP open:"
ObjExcel.Cells(2, col_open_program_titles).Font.Bold 			= TRUE
ObjExcel.Cells(2, col_open_program_status).Value 				= MFIP_active
ObjExcel.Cells(3, col_open_program_titles).Value 				= "SNAP open:"
ObjExcel.Cells(3, col_open_program_titles).Font.Bold 			= TRUE
ObjExcel.Cells(3, col_open_program_status).Value 				= SNAP_active

'If both SNAP and MFIP aren't open, the script will exit
If SNAP_active = False and MFIP_active = False then script_end_procedure("Neither SNAP or MFIP are open. The script will now stop.")




'===================================================================================================================================ASSOCIATING PMIS WITH HH MEMBER NUMBERS
'Now it has to get to STAT/MEMB to associate the HH members with the PMIs
'We do this manually instead of using funclib to maintain the tie to DAIL/DAIL for navigating efficiency while processing many DAILs
EMWriteScreen "stat", 20, 22
EMWriteScreen "memb", 20, 69
If message_type_code = "TIKL" then			'If we're using a TIKL, the month will be all wrong, and it needs to compensate :(
	EMWriteScreen CM_plus_1_mo, 20, 54
	EMWriteScreen CM_plus_1_yr, 20, 57
End if
transmit

'Now we're in STAT/MEMB, and the script will associate each member with their PMI
excel_row = 3 'setting the variable for the following Do...Loop

'Creating headers for the HH member list
ObjExcel.Cells(1, col_HH_memb_PMI_list_memb_num).Value 					= "MEMB/PMI list"
ObjExcel.Cells(1, col_HH_memb_PMI_list_memb_num).Font.Bold 				= True
objExcel.Cells(1, col_HH_memb_PMI_list_memb_num).HorizontalAlignment 	= excel_center
ObjExcel.Range("L1:M1").Merge
ObjExcel.Cells(2, col_HH_memb_PMI_list_memb_num).Value 					= "HH memb"
ObjExcel.Cells(2, col_HH_memb_PMI_list_memb_num).Font.Bold 				= True
ObjExcel.Cells(2, col_HH_memb_PMI_list_PMI).Value 						= "PMI"
ObjExcel.Cells(2, col_HH_memb_PMI_list_PMI).Font.Bold 					= True

'Looping through the panels until it reads each one, which it adds to Excel
Do
	EMReadScreen ref_nbr_on_MEMB, 	2, 4, 33												'Ref nbr = HH memb number
	EMReadScreen PMI_nbr_on_MEMB, 	8, 4, 46												'Reads PMI number on panel
	EMReadScreen current_panel, 	1, 2, 73												'Sees what panel we're on at present
	EMReadScreen amount_of_panels, 	1, 2, 78												'Sees the total number of panels
	PMI_nbr_on_MEMB = Replace(PMI_nbr_on_MEMB, "_", "")										'This allows Ramsey County to use the script. They have underscores here for some reason. Possibly "CAFE"?
	ObjExcel.Cells(excel_row, col_HH_memb_PMI_list_memb_num).Value 	= ref_nbr_on_MEMB		'Adds ref nbr to Excel
	ObjExcel.Cells(excel_row, col_HH_memb_PMI_list_PMI).Value 		= PMI_nbr_on_MEMB		'Adds PMI nbr to Excel
	excel_row = excel_row + 1																'Adds 1 to the Excel row, so that the next panel starts on the next row
	transmit																				'Goes to the next panel
Loop until current_panel = amount_of_panels													'Exits loop once the current panel is the last/only one

'If there's just one memb panel, it means it's a single-individual household, which is not currently a covered option (no logic exists to evaluate it and the policy is murky)
If amount_of_panels = "1" then script_end_procedure("This is a single-individual household, and is not supported by the script. Process manually.")

'Now it's going to use the list of the case's PMIs it just made, and associate a HH member number with each one
'setting the variable for the following Do...Loop
excel_row = 3 			'Resetting this to look at the memb list


Do							'Loops until the HH memb list is out of PMIs
	excel_message_row = 1	'Introducing a new variable which it'll use for comparing the memb side with the message side

	Do						'Loops until the message list is out of messages
		'If...	the PMI from the CSES message equals...						the PMI from the MEMB list...									then the HH member column in the message list...					should equal the ref nbr from the HH memb list.
		If 		ObjExcel.Cells(excel_message_row, col_PMI_number).Value = 	ObjExcel.Cells(excel_row, col_HH_memb_PMI_list_PMI).Value then 	ObjExcel.Cells(excel_message_row, col_HH_memb_number ).Value = 		ObjExcel.Cells(excel_row, col_HH_memb_PMI_list_memb_num).Value

		'Add one more to the excel_message_row so we can check the next message on the next loop
		excel_message_row = excel_message_row + 1
	Loop until ObjExcel.Cells(excel_message_row, col_PMI_number).Value = ""		'Out of messages

	excel_row = excel_row + 1 'Raising the excel row one so that it looks to the next PMI
Loop until ObjExcel.Cells(excel_row, col_HH_memb_PMI_list_PMI).Value = ""		'Out of PMIs!!

'Grabs the case number
call MAXIS_case_number_finder(MAXIS_case_number)

'-----------------------------------------------------------------------------------MOW IT ASSOCIATES EACH MESSAGE WITH A PANEL
'Gets to UNEA
call navigate_to_MAXIS_screen ("STAT", "UNEA")

'Declares the Excel variable for the do...loop which does the checking
excel_row = 3

'Associates panel with each message
Do
	'Creates a UNEA_number variable using the right two characters of a string consisting of "0" and the HH memb column. This prevents issues when running on membs 01-09, which show on Excel as "1-9"
    UNEA_number = right("0" & ObjExcel.Cells(excel_row, col_HH_memb_number).Value, 2)

	'Writes the UNEA_number in the panel and transmits to it (for the first UNEA panel, in case there is more than one)
  	EMWriteScreen UNEA_number, 20, 76
	EMWriteScreen "01", 20, 79
  	transmit

	Do

		'Here's the stuff it needs to read in order to make a decision
		EMReadScreen income_type_on_UNEA, 2, 5, 37			'Reads the income type listed in the panel
		EMReadScreen UNEA_panel_number, 1, 2, 73			'Reads the panel number (a single digit) from the UNEA screen
		EMReadScreen UNEA_panel_total, 1, 2, 78				'Reads the total number of UNEA panels for this member


		'Now it will do things with that info!
    	If income_type_on_UNEA = "__" then 														'First, if it's blank, let's exit the loop entirely so we can move on to the next member. If this member has no panels, anything associated with their member number can be called "NONE". Handling for that is included below.
			UNEA_panel_text = "NONE"															'If there's none for this message it should enter "NONE" so that the user knows to check it
		ElseIf Cint(income_type_on_UNEA) = ObjExcel.Cells(excel_row, col_CS_type).Value then 	'This means it's the correct UNEA panel! And, by extension, the UNEA panel should be documented.
			UNEA_panel_text = "UNEA " & UNEA_number & " 0" & UNEA_panel_number					'The text for the UNEA panel column should be formatted UNEA MM PP, where MM is a two digit UNEA member number and PP is a two digit panel number.
		ElseIf Cint(income_type_on_UNEA) <> ObjExcel.Cells(excel_row, col_CS_type).Value then	'If it's not a UNEA type listed in the message, we should transmit to try again.
			transmit																			'Here's the transmit
			EMReadScreen all_panels_checked, 5, 24, 02											'Checks for the "ENTER A VALID COMMAND" string which occurs when we transmit on the last panel.
    		If all_panels_checked = "ENTER" then UNEA_panel_text = "NONE"						'If that's the case, we can assume this message has no UNEA panel, and we will update those details in the spreadsheet below.
		End if

  	Loop until UNEA_panel_text <> ""			'If it isn't blank, it means we can exit this do...loop (otherwise it should continue to run)

	'Now we need to add the info about the associated panel to the correct column in Excel.

	'This starts with the current Excel row message. We assume this should read the current panel text.
	ObjExcel.Cells(excel_row, col_UNEA_panel).Value = UNEA_panel_text

	'Now we define a new excel_row variable to look at subsequent messages. We want to do this so we don't need to scan every panel more than once. This will allow us to compare the message details below the current one, and to autofill info we just learned. This will save time/transmits.
	excel_row_for_UNEA_panel_autofill = excel_row + 1

	'This loop will compare the HH memb and CS type of the current message with every subsequent message, and will fill info on subsequent messages.
	Do
		'This if...then is for instances where there is no UNEA panel for this member... it'll automatically fill out "NONE" for that memb's other UNEA panels.
		If income_type_on_UNEA = "__" Then
			If ObjExcel.Cells(excel_row_for_UNEA_panel_autofill, col_HH_memb_number).Value = ObjExcel.Cells(excel_row, col_HH_memb_number).Value then ObjExcel.Cells(excel_row, col_UNEA_panel).Value = UNEA_panel_text


		'This catches all other instances, meaning there were UNEA panels
		Else
			If _
			ObjExcel.Cells(excel_row_for_UNEA_panel_autofill, col_HH_memb_number).Value		= ObjExcel.Cells(excel_row, col_HH_memb_number).Value AND _
			ObjExcel.Cells(excel_row_for_UNEA_panel_autofill, col_CS_type).Value 			= ObjExcel.Cells(excel_row, col_CS_type).Value _
			then ObjExcel.Cells(excel_row_for_UNEA_panel_autofill, col_UNEA_panel).Value	= UNEA_panel_text
		End if

		'Increments the row + 1
		excel_row_for_UNEA_panel_autofill = excel_row_for_UNEA_panel_autofill + 1

	Loop until ObjExcel.Cells(excel_row_for_UNEA_panel_autofill, col_msg_number).Value = ""			'Out of messages!



	'Resetting variables
	UNEA_panel_text = ""
	excel_row_for_UNEA_panel_autofill = ""
	income_type_on_UNEA = ""
	excel_row = excel_row + 1



Loop until ObjExcel.Cells(excel_row, col_msg_number).Value = ""			'Loop until we're out of messages

'Now, it makes Excel look prettier (because we all want prettier Excel) by auto-fitting the column width
For col_to_autofit = 1 to col_HH_memb_PMI_list_PMI
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Resetting the variable before the next do...loop
excel_row = 3

'Before we create new data for these UNEA panels, we will copy each bit of message info into an array using a class (MessageDetails).
Do
	total_messages_run = excel_row - 3							'It's minus 2 to account for the two header rows, and minus 1 more because arrays start with 0
	ReDim Preserve message_array(total_messages_run)			'Redims the array to account for the next message, preserves original content
	Set message_array(total_messages_run) = new MessageDetails	'Turns this array item into a new MessageDetails object

	'This grabs each element-to-be-considered-for-the-budget
	message_array(total_messages_run).MsgNum 		= ObjExcel.Cells(excel_row, col_msg_number).Value
	message_array(total_messages_run).PMINum		= ObjExcel.Cells(excel_row, col_PMI_number).Value
	message_array(total_messages_run).MEMBNum		= ObjExcel.Cells(excel_row, col_HH_memb_number).Value
	message_array(total_messages_run).AmtAlloted	= ObjExcel.Cells(excel_row, col_amt_alloted).Value
	message_array(total_messages_run).CSType		= ObjExcel.Cells(excel_row, col_CS_type).Value
	message_array(total_messages_run).IssueDate		= ObjExcel.Cells(excel_row, col_issue_date).Value
	message_array(total_messages_run).UNEAPanel		= ObjExcel.Cells(excel_row, col_UNEA_panel).Value

	'Increments excel_row + 1
	excel_row = excel_row + 1

Loop until ObjExcel.Cells(excel_row, col_msg_number).Value = ""		'Out of messages!

'Creates a blank variable for the next loop
UNEA_panel_array = ""


'Creates an array of just the UNEA Panels
For i = 0 to ubound(message_array)
	UNEA_panel_from_big_array = message_array(i).UNEAPanel
	If InStr (UNEA_panel_array, UNEA_panel_from_big_array) = 0 then UNEA_panel_array = UNEA_panel_array & UNEA_panel_from_big_array & "|"
Next

'UNEA_panel_array has an extra "|" at the end, so this gets rid of it
UNEA_panel_array = left(UNEA_panel_array, len(UNEA_panel_array) - 1 )

'Splits into an array
UNEA_panel_array = split(UNEA_panel_array, "|")

'~~~~~~~~~~~~~~~~~~~~~~~~~~Script check for HRF or REVW due for case noting purposes
'Go to STAT REVW
Call navigate_to_MAXIS_screen ("STAT", "REVW")

'Reads if there is a review that is not received or incomplete - for case noting
EMReadScreen cash_revw_status, 1, 11, 43
EMReadScreen snap_revw_status, 1, 11, 53
EMReadScreen hc_revw_status,   1, 11, 63

If cash_revw_status = "I" or cash_revw_status = "N" Then REVW_due = TRUE
If snap_revw_status = "I" or snap_revw_status = "N" Then REVW_due = TRUE
If hc_revw_status = "I" or hc_revw_status = "N" Then REVW_due = TRUE

'Go to STAT MONT
Call navigate_to_MAXIS_screen ("STAT", "MONT")

'Reads if there is a review that is not received or incomplete - for case noting
EMReadScreen cash_mont_status, 1, 11, 43
EMReadScreen snap_mont_status, 1, 11, 53
EMReadScreen hc_mont_status,   1, 11, 63

If cash_mont_status = "I" or cash_mont_status = "N" Then HRF_Due = TRUE
If snap_mont_status = "I" or snap_mont_status = "N" Then HRF_Due = TRUE
If hc_mont_status = "I" or hc_mont_status = "N" Then HRF_Due = TRUE

'Go back to DAIL'
PF3

'~~~~~~~~~~~~~~~~~~~~~~~Script to determine reporting threshhold
'Navigates to ELIG directly (the DAIL doesn't easily go back to the case-in-question when we use the custom function)
If SNAP_active = TRUE Then
	EMWriteScreen "e", 6, 3
	transmit
	EMWriteScreen "fs", 20, 71
	transmit
	EMWriteScreen "99", 19, 78
	transmit
	row = 17
	Do
		EMReadScreen app_status, 8, row, 50

		If app_status = "APPROVED" Then
			EMReadScreen approval_version, 1, row, 23
			Exit Do
		End If
		row = row - 1
	Loop Until row = 6
	EMWriteScreen approval_version, 18, 54
	transmit
	EMWriteScreen "FSB1", 19, 70
	transmit
	EMReadScreen BUDG_JOBS,	8, 5 , 33
	EMReadScreen BUDG_BUSI,	8, 6 , 33
	EMReadScreen BUDG_PA,   8, 10, 33
	EMReadScreen BUDG_RSDI, 8, 11, 33
	EMReadScreen BUDG_SSI,  8, 12, 33
	EMReadScreen BUDG_VA,   8, 13, 33
	EMReadScreen BUDG_UCWC, 8, 14, 33
	EMReadScreen BUDG_CSES, 8, 15, 33
	EMReadScreen BUDG_OTHR, 8, 16, 33

	If BUDG_JOBS = "        " Then BUDG_JOBS = 0
	If BUDG_BUSI = "        " Then BUDG_BUSI = 0
	If BUDG_PA   = "        " Then BUDG_PA   = 0
	If BUDG_RSDI = "        " Then BUDG_RSDI = 0
	If BUDG_SSI  = "        " Then BUDG_SSI  = 0
	If BUDG_VA   = "        " Then BUDG_VA   = 0
	If BUDG_UCWC = "        " Then BUDG_UCWC = 0
	If BUDG_CSES = "        " Then BUDG_CSES = 0
	If BUDG_OTHR = "        " Then BUDG_OTHR = 0

	EMWriteScreen "FSB2", 19, 70
	transmit

	EMReadScreen FPG_130, 8, 8, 73
	If FPG_130 = "        " THEN FPG_130 = "9999"
	transmit
	EMReadScreen REPT_status, 9, 8, 31
	amount_CS_reported = 0
	CS_Change = FALSE
	Exceed_130 = FALSE

	For i = 0 to ubound(message_array)
		amount_CS_reported = amount_CS_reported + message_array(i).AmtAlloted
	Next

	New_BUDG = abs(BUDG_JOBS) + abs(BUDG_BUSI) + abs(BUDG_PA) + abs(BUDG_RSDI) + abs(BUDG_SSI) + abs(BUDG_VA) + abs(BUDG_UCWC) + abs(BUDG_OTHR) + amount_CS_reported

	If abs(BUDG_CSES) <> amount_CS_reported Then CS_Change = TRUE
	IF New_BUDG >= abs(FPG_130) Then Exceed_130 = TRUE

	'MsgBox ("New budget is: " & New_BUDG & vbNewLine & "CS Reported is: " & amount_CS_reported & vbNewLine & "CS Change: " & CS_Change &vbNewLine & "Exceed 130%: " & Exceed_130)
	PF3
End If

'~~~~~~~~~~~~~~~~~~~~Decision: Is SNAP open? IF YES...
If SNAP_active = true then

	EMWriteScreen "s", 6, 3
	transmit

	'We're going to start by creating a new "MFIP budget" sheet
	ObjExcel.Sheets.Add.Name = "SNAP Budget"

	'Resetting excel_row for this loop
	excel_row = 1

	'This will add info about each UNEA-panel-to-be-changed using a loop
	For each UNEA_panel in UNEA_panel_array

		'If the UNEA_panel isn't "NONE" then it'll add info about the budgets-to-be
		If UNEA_panel <> "NONE" then
			'We'll start with headers consisting of the panel and assorted MFIP details
			ObjExcel.Cells(excel_row + 1, 1).Value = "From the PIC"
			ObjExcel.Cells(excel_row + 2, 1).Value = "Date"
			ObjExcel.Cells(excel_row + 2, 2).Value = "Amt"
			ObjExcel.Cells(excel_row + 2, 4).Value = "Amt"

			'Looking at the new reported Income
			ObjExcel.Cells(excel_row + 1, 6).Value = "CS DAILS Info"
			ObjExcel.Cells(excel_row + 2, 6).Value = "Date"
			ObjExcel.Cells(excel_row + 2, 7).Value = "Amt"

			'Now lets make the fonts bold, because it looks nicer
			ObjExcel.Cells(excel_row	, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 1, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 1, 3).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 2).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 4).Font.Bold = True
			ObjExcel.Cells(excel_row + 1, 6).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 6).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 7).Font.Bold = True

			'Now we'll merge cells for the panel, retro, and pro headers
			ObjExcel.Range("A" & excel_row 		& ":G" & excel_row).Merge
			ObjExcel.Range("A" & excel_row + 1 	& ":D" & excel_row + 1).Merge
			ObjExcel.Range("F" & excel_row + 1 	& ":G" & excel_row + 1).Merge

			'Creates a temporary header_row so that the next loop adds the TYPE info to the correct row
			header_row = excel_row

			'Raises excel_row + 3, so we can start adding messages
			excel_row = excel_row + 3

			'Goes to the correct UNEA panel
			panel_name = left(UNEA_panel, 4)
			person_ref = left(right(UNEA_panel, 5), 2)
			person_ref = right("00" & person_ref, 2)
			version_num = right("00" & right(UNEA_panel, 2), 2)
			EMWriteScreen panel_name, 20, 71
			EMWriteScreen person_ref, 20, 76
			EMWriteScreen version_num, 20, 79
			transmit

			EMReadScreen income_type, 2, 5, 37						'Confirming the type for the spreadshee
			EMWriteScreen "X", 10, 26								'Opens the PIC
			transmit
			EMReadScreen antic_amt, 8, 8, 66						'Gets any amount listed in the right side of the PIC and adds to Excel
			ObjExcel.Cells(excel_row, 4).Value = antic_amt
			ObjExcel.Cells(excel_row, 4).NumberFormat = "$#,##0.00"
			pic_row = 9
			Do 														'Reads all income listed on the left side of PIC
				EMReadScreen date_recvd, 8, pic_row, 13				'Lists it on Excel
				If date_recvd = "__ __ __" Then
					ObjExcel.Cells(excel_row, 1).Value = "-"
					ObjExcel.Cells(excel_row, 2).Value = "-"
					excel_row = excel_row + 1
					Exit Do
				Else
					date_recvd = replace(date_recvd, " ", "/")
					EMReadScreen amt_recvd, 8, pic_row, 25
					ObjExcel.Cells(excel_row, 1).Value = date_recvd
					ObjExcel.Cells(excel_row, 2).Value = amt_recvd
					ObjExcel.Cells(excel_row, 2).NumberFormat = "$#,##0.00"
					excel_row = excel_row + 1
				End If
				pic_row = pic_row + 1
			Loop until pic_row = 14

			EMReadScreen prosp_mo_amt, 8, 18, 56							'Adding what is on the bottom of the PIC
			ObjExcel.Cells(excel_row, 1).Value = "Prosp Monthly Income:"
			ObjExcel.Cells(excel_row	, 1).Font.Bold = True
			ObjExcel.Range("A" & excel_row 		& ":C" & excel_row).Merge

			ObjExcel.Cells(excel_row, 4).Value = prosp_mo_amt
			ObjExcel.Cells(excel_row, 4).NumberFormat = "$#,##0.00"

			'Looks through each message in the array, and if it's for this UNEA panel, it will add it to the Excel sheet
			row_over = header_row + 3
			total_reported = 0
			For i = 0 to ubound(message_array)
				If message_array(i).UNEAPanel <> "NONE" then
					If UNEA_panel = message_array(i).UNEAPanel then
						ObjExcel.Cells(row_over, 6).Value = message_array(i).IssueDate
						ObjExcel.Cells(row_over, 7).Value = message_array(i).AmtAlloted

						ObjExcel.Cells(excel_row, 7).NumberFormat = "$#,##0.00"

						total_reported = total_reported + message_array(i).AmtAlloted
						row_over = row_over + 1

					End if

				End if
			Next

			'Excel headers and formatting
			ObjExcel.Cells(excel_row, 6).Value = "Total:"
			ObjExcel.Cells(excel_row	, 6).Font.Bold = True

			ObjExcel.Cells(excel_row, 7).Value = total_reported
			ObjExcel.Cells(excel_row, 7).NumberFormat = "$#,##0.00"

			ObjExcel.Cells(header_row, 1).Value = UNEA_panel & " | " & "TYPE: " & income_type

			PF3

			excel_row = excel_row + 2		'The next message should start with a row in-between

		Else		'This means all UNEA_panels set to "NONE"

			'If there's a NONE as a UNEA_panel, it'll set this variable to "true", which will force the user to process manually.
			process_manually = true

			'Creates a new excel_row, starting at 3, for messages without panels
			excel_row_no_panel_found = 3

			'Now, if there was no match ("NONE" was listed on the first sheet), it will dump info about those
			For i = 0 to ubound(message_array)

				If message_array(i).UNEAPanel = "NONE" then

					ObjExcel.Cells(1, 6 ).Value = "MESSAGES WITHOUT UNEA PANELS (SPLIT BY HH MEMB)"

					ObjExcel.Cells(2, 6 ).Value = "HH member #"
					ObjExcel.Cells(2, 7 ).Value = "CS type"
					ObjExcel.Cells(2, 8 ).Value = "Amount alloted"
					ObjExcel.Cells(2, 9 ).Value = "Issue date"
					ObjExcel.Cells(2, 10).Value = "Message #"

					ObjExcel.Cells(2, 6 ).Font.Bold	= True
					ObjExcel.Cells(2, 7 ).Font.Bold	= True
					ObjExcel.Cells(2, 8 ).Font.Bold	= True
					ObjExcel.Cells(2, 9 ).Font.Bold	= True
					ObjExcel.Cells(2, 10).Font.Bold	= True



					ObjExcel.Cells(excel_row_no_panel_found, 6 ).Value = "'0" & message_array(i).MEMBNum
					ObjExcel.Cells(excel_row_no_panel_found, 7 ).Value = message_array(i).CSType
					ObjExcel.Cells(excel_row_no_panel_found, 8 ).Value = message_array(i).AmtAlloted
					ObjExcel.Cells(excel_row_no_panel_found, 9 ).Value = message_array(i).IssueDate
					ObjExcel.Cells(excel_row_no_panel_found, 10).Value = message_array(i).MsgNum

					ObjExcel.Cells(excel_row_no_panel_found, 8).NumberFormat = "$#,##0.00"

					excel_row_no_panel_found = excel_row_no_panel_found + 1

				End if
			Next


			ObjExcel.Range("F1:J1").Merge

			ObjExcel.Cells(1, 6).Font.Bold			 = True

		End if

	Next

	'Centering contents
	ObjExcel.Range("A:J").HorizontalAlignment = excel_center

	'Now, it makes Excel look prettier (because we all want prettier Excel) by auto-fitting the column width
	For col_to_autofit = 1 to 10
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	If process_manually = true then
		script_end_procedure("This case appears to be missing a UNEA panel based on the messages received. Evaluate the ''MESSAGES WITHOUT UNEA PANELS'' section of the Excel sheet, add the appropriate panels, and try again or process manually.")
	End If
	PF3

End If

    '~~~~~~~~~~~~~~~~~~~~Displays total and current PIC, user decides if it’s within the realm for each message
close_excel_checkbox = checked 		'Defaulting to have the spreadsheet close after script end

IF SNAP_active = TRUE AND MFIP_active = FALSE THen 	'IF SNAP - NO MFIP

	If Exceed_130 = True OR CS_Change = TRUE Then UNEA_review_checkbox = checked 	'Defaults to have the worker review each panel if income exceeds 130% OR if CS amount is different

	'Dialog defined for if it is a SNAP case.
	BeginDialog CSES_initial_dialog, 0, 0, 296, 140, "CSES Dialog"
	  CheckBox 20, 60, 265, 10, "Check here to review CS UNEA panels for possible adjustments to the budget.", UNEA_review_checkbox
	  EditBox 40, 80, 245, 15, other_notes
	  CheckBox 5, 105, 290, 10, "Check here to have the Excel sheet close at the end of the script run.", close_excel_checkbox
	  EditBox 70, 120, 90, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 185, 120, 50, 15
	    CancelButton 240, 120, 50, 15
	  GroupBox 5, 5, 290, 95, "SNAP Case"
	  Text 190, 15, 95, 10, "SNAP is active on this case"
	  If Exceed_130 = True Then Text 10, 25, 205, 10, "It appears that the income for this case may exceed 130% FPG."
	  IF CS_Change = TRUE Then Text 10, 35, 275, 10, "It appears there is a difference between the budgeted CS Income and DAIL Message Income."
	  If REVW_due = TRUE Then Text 10, 45, 200, 10, "There is a review due that has not been recevied/processed."
	  Text 10, 85, 25, 10, "Notes"
	  Text 5, 125, 65, 10, "Worker signature:"
	EndDialog

ElseIf MFIP_active = TRUE Then

	'Dialog specific to MFIP cases
	BeginDialog CSES_initial_dialog, 0, 0, 296, 120, "CSES Dialog"
	  EditBox 40, 60, 245, 15, other_notes
	  CheckBox 5, 85, 290, 10, "Check here to have the Excel sheet close at the end of the script run.", close_excel_checkbox
	  EditBox 70, 100, 90, 15, worker_signature
	  ButtonGroup ButtonPressed
	    OkButton 185, 100, 50, 15
	    CancelButton 240, 100, 50, 15
	  GroupBox 5, 5, 285, 75, "MFIP Case"
	  Text 190, 15, 95, 10, "MFIP is active on this case."
	  If HRF_due = TRUE Then Text 10, 25, 195, 10, "Case has a HRF Due that has not been received/processed."
	  If REVW_due = TRUE Then Text 10, 35, 205, 10, "Case has a Review due that has not been received/processed."
	  Text 10, 45, 220, 10, "Script will review case for UNEA to be update with CS Income from DAILs."
	  Text 10, 65, 25, 10, "Notes"
	  Text 5, 105, 65, 10, "Worker signature:"
	EndDialog

End IF

'Runs the dialog from above
Do
	err_msg = ""
	Dialog CSES_initial_dialog
	cancel_confirmation
	If worker_signature = "" Then err_msg = err_msg & "Please sign your case note"
	If err_msg <> "" Then MsgBox (err_msg)
Loop until err_msg = ""

'If the worker signature is the Konami code (UUDDLRLRBA), developer mode will be triggered
If worker_signature = "UUDDLRLRBA" then
    MsgBox "Developer mode: ACTIVATED!"
    developer_mode = true           'This will be helpful later
    collecting_statistics = false   'Lets not collect this, shall we?		'<<<<CHECK THIS, I THINK THE VARIABLE IS CALLED SOMETHING DIFFERENT IN THE FUNCTION
End if

'Each UNEA panel will be reviewed individually for cases that may need adjustment
If SNAP_active = true AND UNEA_review_checkbox = checked then
	Dim PIC_Payment_array()
	ReDim PIC_Payment_array(5)
	EMWriteScreen "s", 6, 3
	transmit
	counter = 0

	'This will add info about each UNEA-panel-to-be-changed using a loop
	For each UNEA_panel in UNEA_panel_array
		'Go to the correct panel
		panel_name = left(UNEA_panel, 4)
		person_ref = left(right(UNEA_panel, 5), 2)
		person_ref = right("00" & person_ref, 2)
		version_num = right("00" & right(UNEA_panel, 2), 2)
		EMWriteScreen panel_name, 20, 71
		EMWriteScreen person_ref, 20, 76
		EMWriteScreen version_num, 20, 79
		transmit

		'Gather all the data for the dialog display
		EMReadScreen income_type, 2, 5, 37
		EMWriteScreen "X", 10, 26
		transmit
		EMReadScreen date_of_calc, 8, 5, 34
		date_of_calc = replace(date_of_calc, " ", "/")
		EMReadScreen antic_amt, 8, 8, 66
		EMReadScreen pay_freq, 1, 5, 64
		If pay_freq = "1" Then pay_freq = "1 - MONTHLY"
		If pay_freq = "2" Then pay_freq = "2 - 2X/MONTH"
		If pay_freq = "3" Then pay_freq = "3 - BIWEEKLY"
		If pay_freq = "4" Then pay_freq = "4 - WEEKLY"
		EMReadScreen reg_non_mo, 8, 12, 66
		EMReadScreen num_of_mo, 2, 13, 64
		EMReadScreen average_income, 8, 17, 56
		EMReadScreen prosp_mo_amt, 8, 18, 56
		EMReadScreen total_recvd, 8, 14, 25

		ReDim PIC_Payment_array(5)
		total_to_count = 0

		For payment = 0 to 4
			EMReadScreen date_recvd, 8, payment + 9, 13
			If date_recvd = "__ __ __" Then
				PIC_Payment_array(payment) = "__ __ __    ________"
			Else
				date_recvd = replace(date_recvd, " ", "/")
				EMReadScreen amt_recvd, 8, payment + 9, 25
				PIC_Payment_array(payment) = date_recvd & "   $" & amt_recvd
			End If
		Next
		x = 0

		'Dynamic Dialog that will mimmick the PIC and ask for worker Input
		BeginDialog pic_review_dialog, 0, 0, 346, 170, "REVIEW THE PIC"
		  Text 170, 5, 50, 10, UNEA_panel
		  GroupBox 5, 15, 225, 130, "PIC"
		  Text 10, 25, 65, 10, "Date of Calculation:"
		  Text 80, 25, 30, 10, date_of_calc
		  Text 130, 25, 35, 10, "Pay Freq."
		  Text 170, 25, 55, 10, pay_freq
		  Text 20, 35, 60, 10, "Income Received"
		  For line = 0 to 4
		  	Text 15, 45 + line * 10, 75, 10, PIC_Payment_array(line)
		  Next
		  Text 25, 100, 20, 10, "Total:"
		  Text 55, 100, 35, 10, "$" & total_recvd
		  Text 130, 40, 65, 10, "Anticipated Income:"
		  Text 165, 50, 15, 10, "Amt:"
		  Text 190, 50, 30, 10, "$" & antic_amt
		  Text 130, 70, 75, 10, "Regular Non-Monthly:"
		  Text 165, 85, 15, 10, "Amt:"
		  Text 190, 85, 35, 10, "$" & reg_non_mo
		  Text 130, 95, 65, 10, "Number of Months"
		  Text 200, 95, 5, 10, num_of_mo
		  Text 145, 120, 50, 10, "$" & average_income
		  Text 145, 130, 50, 10, "$" & prosp_mo_amt
		  Text 65, 130, 75, 10, "Prosp Monthly Income:"
		  GroupBox 240, 15, 95, 130, "CS Messages Recevied"
		  Text 75, 120, 65, 10, "Average /Pay Date:"
		  For i = 0 to ubound(message_array)
			  If message_array(i).UNEAPanel <> "NONE" then
				  If UNEA_panel = message_array(i).UNEAPanel then
					  Text 250, 30 + 10 * x, 45, 10, message_array(i).IssueDate
					  Text 290, 30 + 10 * x, 50, 10, "$" & message_array(i).AmtAlloted
					  total_to_count = total_to_count + message_array(i).AmtAlloted
					  x = x + 1
				  End if
			  End if
		  Next
		  Text 245, 130, 25, 10, "Total:"
		  Text 280, 130, 50, 10, "$" & total_to_count
		  Text 30, 155, 205, 10, "Does this income require rebudgeting and a new approval?"
		  ButtonGroup ButtonPressed
		    PushButton 245, 150, 25, 15, "YES", yes_button
		    PushButton 275, 150, 25, 15, "NO", No_button
		EndDialog

		'The only options on this dialog are Yes or No
		Dialog pic_review_dialog
		If ButtonPressed = yes_button Then 			'If worker clicks 'yes' then the income needs to be rebudgeted and is 'out of the realm'
			UNEA_panel_array(counter) = UNEA_panel_array(counter) & " YES"
			Outside_the_realm = TRUE
		End If 										'If worker clicks'no' they are confirming that the income is within acceptable range
		If ButtonPressed = no_button  Then UNEA_panel_array(counter) = UNEA_panel_array(counter) & "  NO"

		counter = counter + 1		'This is to add detail to the array information so that in future enhancements we can update the UNEA panels
		EMReadScreen PIC_check, 35, 3, 28
		If PIC_check = "SNAP Prospective Income Calculation" Then PF3
	Next

End If

'~~~~~~~~~~~~~~~~~~~~Decision: Is MFIP/DWP open? IF YES...
If MFIP_active = true then

	'Navigates to STAT directly (the DAIL doesn't easily go back to the case-in-question when we use the custom function)
	EMWriteScreen "s", 6, 3
	transmit

	'We're going to start by creating a new "MFIP budget" sheet
	ObjExcel.Sheets.Add.Name = "MFIP Budget"

	'Resetting excel_row for this loop
	excel_row = 1

	'This will add info about each UNEA-panel-to-be-changed using a loop
	For each UNEA_panel in UNEA_panel_array

		'If the UNEA_panel isn't "NONE" then it'll add info about the budgets-to-be
		If UNEA_panel <> "NONE" then
			'We'll start with headers consisting of the panel and assorted MFIP details
			'ObjExcel.Cells(excel_row	, 1).Value = UNEA_panel
			ObjExcel.Cells(excel_row + 1, 1).Value = "Retrospective"
			ObjExcel.Cells(excel_row + 1, 3).Value = "Prospective"
			ObjExcel.Cells(excel_row + 2, 1).Value = "Date"
			ObjExcel.Cells(excel_row + 2, 2).Value = "Amt"
			ObjExcel.Cells(excel_row + 2, 3).Value = "Date"
			ObjExcel.Cells(excel_row + 2, 4).Value = "Amt"

			'Now lets make the fonts bold, because it looks nicer
			ObjExcel.Cells(excel_row	, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 1, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 1, 3).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 1).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 2).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 3).Font.Bold = True
			ObjExcel.Cells(excel_row + 2, 4).Font.Bold = True

			'Now we'll merge cells for the panel, retro, and pro headers
			'excel_header_cols = "A" & excel_row & ":D" & excel_row
			ObjExcel.Range("A" & excel_row 		& ":D" & excel_row).Merge
			ObjExcel.Range("A" & excel_row + 1 	& ":B" & excel_row + 1).Merge
			ObjExcel.Range("C" & excel_row + 1 	& ":D" & excel_row + 1).Merge

			'Creates a temporary header_row so that the next loop adds the TYPE info to the correct row
			header_row = excel_row

			'Raises excel_row + 3, so we can start adding messages
			excel_row = excel_row + 3

			'Looks through each message in the array, and if it's for this UNEA panel, it will add it to the Excel sheet
			For i = 0 to ubound(message_array)
				If message_array(i).UNEAPanel <> "NONE" then
					If UNEA_panel = message_array(i).UNEAPanel then
						ObjExcel.Cells(excel_row, 1).Value = message_array(i).IssueDate
						ObjExcel.Cells(excel_row, 2).Value = message_array(i).AmtAlloted
						ObjExcel.Cells(excel_row, 3).Value = dateadd("m", 2, message_array(i).IssueDate)
						ObjExcel.Cells(excel_row, 4).Value = message_array(i).AmtAlloted

						ObjExcel.Cells(excel_row, 2).NumberFormat = "$#,##0.00"
						ObjExcel.Cells(excel_row, 4).NumberFormat = "$#,##0.00"

						ObjExcel.Cells(header_row, 1).Value = UNEA_panel & " | " & "TYPE: " & message_array(i).CSType

						excel_row = excel_row + 1

					End if

				End if
			Next

			excel_row = excel_row + 1		'The next message should start with a row in-between

		Else		'This means all UNEA_panels set to "NONE"

			'If there's a NONE as a UNEA_panel, it'll set this variable to "true", which will force the user to process manually.
			process_manually = true

			'Creates a new excel_row, starting at 3, for messages without panels
			excel_row_no_panel_found = 3

			'Now, if there was no match ("NONE" was listed on the first sheet), it will dump info about those
			For i = 0 to ubound(message_array)

				If message_array(i).UNEAPanel = "NONE" then

					ObjExcel.Cells(1, 6 ).Value = "MESSAGES WITHOUT UNEA PANELS (SPLIT BY HH MEMB)"

					ObjExcel.Cells(2, 6 ).Value = "HH member #"
					ObjExcel.Cells(2, 7 ).Value = "CS type"
					ObjExcel.Cells(2, 8 ).Value = "Amount alloted"
					ObjExcel.Cells(2, 9 ).Value = "Issue date"
					ObjExcel.Cells(2, 10).Value = "Message #"

					ObjExcel.Cells(2, 6 ).Font.Bold	= True
					ObjExcel.Cells(2, 7 ).Font.Bold	= True
					ObjExcel.Cells(2, 8 ).Font.Bold	= True
					ObjExcel.Cells(2, 9 ).Font.Bold	= True
					ObjExcel.Cells(2, 10).Font.Bold	= True



					ObjExcel.Cells(excel_row_no_panel_found, 6 ).Value = "'0" & message_array(i).MEMBNum
					ObjExcel.Cells(excel_row_no_panel_found, 7 ).Value = message_array(i).CSType
					ObjExcel.Cells(excel_row_no_panel_found, 8 ).Value = message_array(i).AmtAlloted
					ObjExcel.Cells(excel_row_no_panel_found, 9 ).Value = message_array(i).IssueDate
					ObjExcel.Cells(excel_row_no_panel_found, 10).Value = message_array(i).MsgNum

					ObjExcel.Cells(excel_row_no_panel_found, 8).NumberFormat = "$#,##0.00"

					excel_row_no_panel_found = excel_row_no_panel_found + 1

				End if
			Next

			'Excel Formatting
			ObjExcel.Range("F1:J1").Merge
			ObjExcel.Cells(1, 6).Font.Bold			 = True
		End if
	Next

	'Centering contents
	ObjExcel.Range("A:J").HorizontalAlignment = excel_center

	'Now, it makes Excel look prettier (because we all want prettier Excel) by auto-fitting the column width
	For col_to_autofit = 1 to 10
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	If process_manually = true then
		script_end_procedure("This case appears to be missing a UNEA panel based on the messages received. Evaluate the ''MESSAGES WITHOUT UNEA PANELS'' section of the Excel sheet, add the appropriate panels, and try again or process manually.")
	Else
		MFIP_evaluation_popup = MsgBox("The planned budget is indicated on the Excel spreadsheet. Evaluate the budget as indicated. If the budget appears correct for each UNEA panel indicated, press OK to update MAXIS and case note. Otherwise, press cancel and process manually.", vbOKCancel)
		If MFIP_evaluation_popup = vbCancel then script_end_procedure("Script ended due to MFIP budget being marked as incorrect. The script will now close.")
		'The script will now update the MFIP budget if NOT developer mode
		If developer_mode <> TRUE Then
			'First, make sure we are in correct month to update the panel (issue month + 2)
			UNEA_month = dateadd("M", 2, issue_date)
			IF len(UNEA_month) = 1 THEN UNEA_month = "0" & UNEA_month
			MAXIS_footer_month = UNEA_month
			back_to_self 'We go back to the self menu, because navigate_to_maxis_screen doesn't update footer month when already in STAT
			Call navigate_to_MAXIS_screen ("STAT", "UNEA")
			'ObjExcel.Sheets("Message and Disb info").Activate
			row = 1

			Do
				If left(ObjExcel.Cells(row, 1).Value, 4) = "UNEA" Then
					'Navigates to the correct UNEA panel
					panel_to_update = left(ObjExcel.Cells(row, 1).Value, 10)
					memb_to_update = right(left(panel_to_update, 7), 2)
					panel_numb_to_update = right(panel_to_update, 2)
					EMWriteScreen panel_to_update, 20, 71
					EMWriteScreen memb_to_update, 20, 76
					EMWriteScreen panel_numb_to_update, 20, 79
					transmit
					PF9
					'Now it updates the code to be a "6" for verification type
					EMWriteScreen "6", 5, 65
					'This bit will erase all of the previous data that was listed on UNEA
					For unea_row = 13 to 17		'Needs to go through each row of income data
						unea_col = 25
						Do
							EMSetCursor unea_row, unea_col		'This is basically pressing 'end' on that field in UNEA
							EMSendKey "<eraseeof>"
							unea_col = unea_col + 3				'the date fields are 3 apart
							If unea_col = 34 Then unea_col = 39	'income field
							If unea_col = 42 Then unea_col = 54	'prospective side
							If unea_col = 63 Then unea_col = 68	'income field
						Loop until unea_col = 71
					Next
					row = row + 1
					unea_row = 13
					Do
						If IsDate(ObjExcel.Cells(row, 1).Value) = TRUE Then
							Call create_mainframe_friendly_date (ObjExcel.Cells(row, 1).Value, unea_row, 25, "YY")
							retro_amt = FormatNumber(ObjExcel.Cells(row, 2).Value, 2, , , 0)
							EMWriteScreen retro_amt, unea_row, 39
							Call create_mainframe_friendly_date (ObjExcel.Cells(row, 3).Value, unea_row, 54, "YY")
							prosp_amt = FormatNumber(ObjExcel.Cells(row, 4).Value, 2, , , 0)
							EMWriteScreen prosp_amt, unea_row, 68
							unea_row = unea_row + 1
						End If

						row = row + 1
					Loop until ObjExcel.Cells(row, 1).Value = ""
					transmit
					transmit
				End If
				row = row + 1
			Loop until ObjExcel.Cells(row, 1).Value = ""
		End If
	End if

End if
	

'Alert to worker that additional action is required.
If Outside_the_realm = TRUE Then MsgBox "This is a SNAP case and you have indicated at least one of the UNEA panels needs to be reviewed for possible budget adjustment." & vbNewLine & vbNewLine & "At this time, this script does NOT update UNEA for SNAP cases. Case note will indicate that worker followup is needed."

    '~~~~~~~~~~~~~~~~~~~~Case note details from Excel sheet, and all panels updated
'Navigates to CASE/NOTE directly (the DAIL doesn't easily go back to the case-in-question when we use the custom function)
If developer_mode <> TRUE Then
	'Check to make sure we are back to our dail
	EMReadScreen DAIL_check, 4, 2, 48
	IF DAIL_check <> "DAIL" THEN
		PF3 'This should bring us back from UNEA or other screens
		EMReadScreen DAIL_check, 4, 2, 48
		IF DAIL_check <> "DAIL" THEN 'If we are still not at the dail, try to get there using custom function, this should result in being on the correct dail (but not 100%)
			call navigate_to_MAXIS_screen("DAIL", "DAIL")
		END IF
	END IF
	EMWriteScreen "n", 6, 3
	transmit

	PF9
	EMReadScreen case_note_mode_check, 7, 20, 3
	If case_note_mode_check <> "Mode: A" then MsgBox "You are not in a case note on edit mode. You might be in inquiry. Try the script again in production."
	If case_note_mode_check <> "Mode: A" then end_excel_and_script

	If REVW_due = TRUE Then
		Call Write_Variable_in_CASE_NOTE (":::CSES Messages Reviewed:::: REVIEW DUE")
	ElseIf HRF_Due = TRUE Then
		Call Write_Variable_in_CASE_NOTE (":::CSES Messages Reviewed:::: HRF DUE")
	Else
		Call Write_Variable_in_CASE_NOTE (":::CSES Messages Reviewed::::")
	End If
	Call Write_Variable_in_CASE_NOTE ("* Income reported from PRISM Interface - details are listed in previous case notes.")
	If MFIP_active = TRUE Then Call Write_Variable_in_CASE_NOTE ("* Updated retro/prospective income amounts.")
	If Exceed_130 = TRUE Then Call Write_Variable_in_CASE_NOTE ("* With this CS Income, it appears case income may exceed 130% FPG.")
	If CS_Change = TRUE Then
		Call Write_Variable_in_CASE_NOTE ("* CS Income listed in DAILs is different from the amount of CS Income Budgeted.")
		Call Write_Bullet_and_Variable_in_Case_Note ("CS Income Budgeted", BUDG_CSES)
		Call Write_Bullet_and_Variable_in_Case_Note ("CS Income From DAIL", amount_CS_reported)
	End If

'reading from excel sheet
IF SNAP_active = TRUE Then
	Dim xlApp 
	Dim xlBook 
	Dim xlSheet 
	RowCN = 1
	Set objSheet = objExcel.ActiveWorkbook.Worksheets("SNAP Budget") 
	Do
		MEMBandTYPE = Trim(objSheet.Cells(RowCN, 1).Value)
		Do
			RowCN = RowCN + 1
			rowCHECK = Trim(objSheet.Cells(RowCN, 1).Value)	
			IF rowCHECK = "Prosp Monthly Income:" then amount = Trim(objSheet.Cells(RowCN, 7).Value)
		Loop until rowCHECK = "Prosp Monthly Income:"
		RowCN = RowCN + 2
		Call Write_Variable_in_Case_Note ("     " & MEMBandTYPE & " - $" & amount)
		blankCHECK = Trim(objSheet.Cells(RowCN, 1).Value)	
	Loop Until blankCHECK = ""
End IF

	If Outside_the_realm <> TRUE AND UNEA_review_checkbox = checked Then Call Write_Variable_in_CASE_NOTE ("* FS PIC reviewed, adjustments to budget not needed.")
	If Outside_the_realm <> TRUE AND UNEA_review_checkbox = unchecked Then Call Write_Variable_in_CASE_NOTE ("* FS Budget reviewed, adjustments to budget not needed.")
	If Outside_the_realm = TRUE Then Call Write_Variable_in_CASE_NOTE ("* FS PIC Reviewed, update needed - worker to process manually.")
	IF MFIP_active = TRUE  AND FS_active = TRUE Then Call Write_Variable_in_CASE_NOTE ("* FS PIC not evaluated, as case also has MFIP.")
	Call Write_Bullet_and_Variable_in_Case_Note ("Notes", other_notes)

	Call Write_Variable_in_CASE_NOTE("---")
	Call Write_Variable_in_CASE_NOTE(worker_signature & ", using automated script")
Else
'Developer mode shows a message box of the case note
	Case_Note_Message = ""
	If REVW_due = TRUE Then
		Case_Note_Message = Case_Note_Message & vbNewLine & ":::CSES Messages Reviewed:::: REVIEW DUE"
	ElseIf HRF_Due = TRUE Then
		Case_Note_Message = Case_Note_Message & vbNewLine & ":::CSES Messages Reviewed:::: HRF DUE"
	Else
		Case_Note_Message = Case_Note_Message & vbNewLine & ":::CSES Messages Reviewed::::"
	End If
	Case_Note_Message = Case_Note_Message & vbNewLine & "* Income reported from PRISM Interface - details are listed in previous case notes."
	If MFIP_active = TRUE Then Case_Note_Message = Case_Note_Message & vbNewLine & "* Updated retro/prospective income amounts."
	If Exceed_130 = TRUE Then Case_Note_Message = Case_Note_Message & vbNewLine & "* With this CS Income, it appears case income may exceed 130% FPG."
	If CS_Change = TRUE Then
		Case_Note_Message = Case_Note_Message & vbNewLine & "* CS Income listed in DAILs is different from the amount of CS Income Budgeted."
		Case_Note_Message = Case_Note_Message & vbNewLine & "CS Income Budgeted: " &  BUDG_CSES
		Case_Note_Message = Case_Note_Message & vbNewLine & "CS Income From DAIL:" & amount_CS_reported
	End If
	If Outside_the_realm <> TRUE AND UNEA_review_checkbox = checked Then Case_Note_Message = Case_Note_Message & vbNewLine &"* FS PIC reviewed, adjustments to budget not needed."
	If Outside_the_realm <> TRUE AND UNEA_review_checkbox = unchecked Then Case_Note_Message = Case_Note_Message & vbNewLine &"* FS Budget reviewed, adjustments to budget not needed."
	If Outside_the_realm = TRUE Then Case_Note_Message = Case_Note_Message & vbNewLine & "* FS PIC Reviewed, update needed - worker to process manually."
	IF MFIP_active = TRUE  AND FS_active = TRUE Then Case_Note_Message = Case_Note_Message & vbNewLine & "* FS PIC not evaluated, as case also has MFIP."
	Case_Note_Message = Case_Note_Message & vbNewLine & "Notes:" & other_notes

	Case_Note_Message = Case_Note_Message & vbNewLine & "---"
	Case_Note_Message = Case_Note_Message & vbNewLine & worker_signature & ", using automated script"

	MsgBox Case_Note_Message
End If

If close_excel_checkbox = checked Then
	'Manually closing workbooks so that the stats script can finish up
	objExcel.DisplayAlerts = False
	objExcel.Workbooks.Close
	objExcel.quit
	objExcel.DisplayAlerts = True
End If

script_end_procedure("")
