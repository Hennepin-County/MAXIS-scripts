excel_visible_checkbox = 1		'<<<<<<<<<<<<<<<<<<<THIS IS TEMPORARY, JUST FOR TESTING
run_locally = true				'<<<<<<<<<<<<<<<<<<<THIS IS TEMPORARY, JUST FOR TESTING

'GATHERING STATS===========================================================================================================
name_of_script = "DAIL - CSES SCRUBBER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY==========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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
'END FUNCTIONS LIBRARY BLOCK===============================================================================================

'Required for statistical purposes=========================================================================================
STATS_counter = 0              'sets the stats counter at 0 because each iteration of the loop which counts the dail messages adds 1 to the counter.
STATS_manualtime = 54          'manual run time in seconds
STATS_denomination = "I"       'I is for each dail message
'END OF stats block========================================================================================================

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

'Displays initial dialog (script assumes you're on a CSES message by virtue of the DAIL scrubber). Dialog has no mandatory fields, so there's no loop.
'Dialog CSES_initial_dialog			<<<<RESET THIS PLEEEEEEEEEEEEEEEEEEEEEEEEEEEASE
'If ButtonPressed = cancel then stopscript	<<<<RESET THIS PLEEEEEEEEEEEEEEEEEEEEEEEEEEEASE

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
		'<<<<<<<<<<<<<<<<<<<<PROBABLY WHERE PENNY ISSUE SHOULD GO, MAYBE JUST ADD PARTIALS TO THE FIRST MEMB????????
    	ObjExcel.Cells(excel_row, col_CS_type).Value 		= CS_type						'This is the type, and it's helpful to know this when we write to UNEA
    	ObjExcel.Cells(excel_row, col_issue_date).Value 	= issue_date						'The date it was issued
    	excel_row = excel_row + 1											'Increments up one in order to start on the next Excel row
    Next
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
If row <> 0 then MsgBox "As of March 2016 the health care sections have been removed from the CSES Scrubber. Evaluate any health care ramifications manually."

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
call MAXIS_case_number_finder(case_number)

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
	total_messages_run = excel_row - 2							'It's minus 2 to account for the two header rows
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


'~~~~~~~~~~~~~~~~~~~~Decision: Is MFIP/DWP open? IF YES...
If MFIP_active = true then

	'We're going to start by creating a new "MFIP budget" sheet
	ObjExcel.Sheets.Add.Name = "MFIP Budget"
	
	'Resetting excel_row for this loop
	excel_row = 1
	
	'This will add info about each UNEA-panel-to-be-changed using a loop
	'Do
		'We'll start with a header consisting of the panel and its type
		'ObjExcel.Cells(excel_row, 1).Value = 
	
	
End if

    '~~~~~~~~~~~~~~~~~~~~Displays prospective estimated budget based on DAILs received

    '~~~~~~~~~~~~~~~~~~~~Decision: Does user want to update? IF YES...

        '~~~~~~~~~~~~~~~~~~~~Script updates UNEA for all messages with prospective amounts and actual amounts for retrospective budgeting

'~~~~~~~~~~~~~~~~~~~~Decision: Is SNAP open? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Displays total and current PIC, user decides if it’s within the realm for each message

'~~~~~~~~~~~~~~~~~~~~Decision: Any panels updated/checked for either SNAP or MFIP? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Case note details from Excel sheet, and all panels updated

'~~~~~~~~~~~~~~~~~~~~Decision: Does user want Excel breakdown of info? IF YES...

    '~~~~~~~~~~~~~~~~~~~~Make Excel visible


script_end_procedure("")
