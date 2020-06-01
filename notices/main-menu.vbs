'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MAIN MENU.vbs"
start_time = timer

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
call changelog_update("06/01/2020", "Temporary removal of DWP ES REFERRAL script. DWP referrals are currently not being offered.", "Ilse Ferris, Hennepin County")
call changelog_update("03/26/2019", "Retired 'NOTICES - MNSURE MEMO'. Verbiage has been updated and added to the 'NOTES - METS to MAXIS TRANSITION' script.", "Ilse Ferris, Hennepin County")
call changelog_update("09/30/2018", "Consolidated WCOMs to one script. New submenus of script categories.", "Casey Love, Hennepin County")
call changelog_update("03/01/2018", "Removed NOTICES scripts APPOINTMENT LETTER and NOMI. This process has been automated through the On Demand Waiver process.", "Ilse Ferris, Hennepin County")
call changelog_update("09/25/2017", "Added new script: SNAP WCOM - Failure to Comply WCOM.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_NOTICES_menu_dialog(script_array, sub_menu)
	BeginDialog NOTICES_dialog, 0, 0, 516, 340, "NOTICES Scripts"
	 	Text 5, 5, 435, 10, "Notices scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 40, 15, "NOTICES", 				NOTICES_main_button
            PushButton 055, 35, 60, 15, "HEALTH CARE",          HEALTH_CARE_button
            PushButton 115, 35, 35, 15, "SNAP",                 SNAP_button
            PushButton 150, 35, 35, 15, "CASH",                 CASH_button
            PushButton 185, 35, 55, 15, "WORD DOCS",            WORD_DOCS_button
            PushButton 260, 35, 70, 15, "ADD WCOM",             ADD_WCOM_button
		 	'PushButton 055, 55, 60, 15, "SNAP WCOMS", 			SNAP_WCOMS_button
		 	PushButton 445, 10, 65, 10, "Instructions", 		Instructions_button

		'This starts here, but it shouldn't end here :)
		vert_button_position = 70

		For current_script = 0 to ubound(script_array)
            If InStr(script_array(current_script).subcategory, sub_menu) <> 0 Then
    			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
    			'FUNCTION		HORIZ. ITEM POSITION								VERT. ITEM POSITION		ITEM WIDTH									ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
    			PushButton 		5, 													vert_button_position, 	script_array(current_script).button_size, 	10, 			script_array(current_script).script_name, 			button_placeholder
    			Text 			script_array(current_script).button_size + 10, 		vert_button_position, 	500, 										10, 			"--- " & script_array(current_script).description
    			'----------
    			vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
    			'----------
    			script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
    			button_placeholder = button_placeholder + 1
            End If
		next

		CancelButton 460, 320, 50, 15
		GroupBox 5, 20, 240, 35, "NOTICES Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM Instructions_button
dim NOTICES_dialog

script_array_NOTICES_main = array()
script_array_NOTICES_list = array()


'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------NOTICES MAIN MENU

'Resetting the variable
script_num = 0												'establishes value of scripts at 0
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name 			= "12 Month Contact"																'Script name
script_array_NOTICES_main(script_num).file_name 			= "12-month-contact.vbs"															'Script URL
script_array_NOTICES_main(script_num).description 			= "Sends a MEMO to the client of their reporting responsibilities (required for SNAP 2-yr certifications, per POLI/TEMP TE02.08.165)."
script_array_NOTICES_main(script_num).subcategory           = "SNAP"

'script_num = script_num + 1									'Increment by one
'ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
'Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
'script_array_NOTICES_main(script_num).script_name 			= "DWP ES Referral"															'Script name
'script_array_NOTICES_main(script_num).file_name 			= "dwp-es-referral.vbs"														'Script URL
'script_array_NOTICES_main(script_num).description 			= "Creates a case note, a manual referral in INFC/WF1M and sends a SPEC/MEMO to the client."
'script_array_NOTICES_main(script_num).subcategory           = "CASH"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name 			= "Eligibility Notifier"															'Script name
script_array_NOTICES_main(script_num).file_name 			= "eligibility-notifier.vbs"														'Script URL
script_array_NOTICES_main(script_num).description 			= "Sends a MEMO informing client of possible program eligibility for SNAP, MA, MSP, MNsure or CASH."
script_array_NOTICES_main(script_num).subcategory           = "SNAP~HC~CASH"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "LTC Asset Transfer"
script_array_NOTICES_main(script_num).file_name				= "ltc-asset-transfer.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO to a LTC client regarding asset transfers. "
script_array_NOTICES_main(script_num).subcategory           = "HC"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "MA Inmate Application WCOM"
script_array_NOTICES_main(script_num).file_name				= "ma-inmate-application-wcom.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a WCOM on a MA notice for Inmate Applications"
script_array_NOTICES_main(script_num).subcategory           = "HC"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "MA-EPD No Initial Premium Paid"
script_array_NOTICES_main(script_num).file_name				= "ma-epd-no-initial-premium.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a WCOM on a denial for no initial MA-EPD premium."
script_array_NOTICES_main(script_num).subcategory           = "HC"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name 			= " MEMO to Word "
script_array_NOTICES_main(script_num).file_name				= "memo-to-word.vbs"
script_array_NOTICES_main(script_num).description 			= "Copies a MEMO or WCOM from MAXIS and formats it in a Word Document."
script_array_NOTICES_main(script_num).subcategory           = "WORD"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= " Method B WCOM "													'needs spaces to generate button width properly.
script_array_NOTICES_main(script_num).file_name				= "method-b-wcom.vbs"
script_array_NOTICES_main(script_num).description			= "Makes detailed WCOM regarding spenddown vs. recipient amount for method B HC cases."
script_array_NOTICES_main(script_num).subcategory           = "HC"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "Out Of State"
script_array_NOTICES_main(script_num).file_name				= "out-of-state.vbs"
script_array_NOTICES_main(script_num).description			= "Generates out of state inquiry-Microsoft word document notice that can be use to fax."
script_array_NOTICES_main(script_num).subcategory           = "WORD"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "PA Verif Request"
script_array_NOTICES_main(script_num).file_name				= "pa-verif-request.vbs"
script_array_NOTICES_main(script_num).description			= "Creates a Word document with PA benefit totals for other agencies to determine client benefits."
script_array_NOTICES_main(script_num).subcategory           = "WORD"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name 			= "Resources Notifier"															'Script name
script_array_NOTICES_main(script_num).file_name 			= "resources-notifier.vbs"														'Script URL
script_array_NOTICES_main(script_num).description 			= "Sends a MEMO informing client of some possible outside resources."
script_array_NOTICES_main(script_num).subcategory           = "SNAP~HC~CASH~WORD"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "SNAP E and T Letter"
script_array_NOTICES_main(script_num).file_name				= "snap-e-and-t-letter.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a SPEC/LETR informing client that they have an Employment and Training appointment."
script_array_NOTICES_main(script_num).subcategory           = "SNAP"

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "Verifications Still Needed"
script_array_NOTICES_main(script_num).file_name				= "verifications-still-needed.vbs"
script_array_NOTICES_main(script_num).description			= "Creates a Word document informing client of a list of verifications that are still required."
script_array_NOTICES_main(script_num).subcategory           = "WORD"

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
NOTICES_main_button		= 1000
'SNAP_WCOMS_button		= 2000
HEALTH_CARE_button      = 2000
SNAP_button             = 3000
CASH_button             = 4000
WORD_DOCS_button        = 5000
ADD_WCOM_button         = 6000

'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = NOTICES_main_button then
		Call declare_NOTICES_menu_dialog(script_array_NOTICES_main, "")
	' ElseIf ButtonPressed = SNAP_WCOMS_button then
	' 	declare_NOTICES_menu_dialog(script_array_NOTICES_list)
    ElseIf ButtonPressed = HEALTH_CARE_button then
        Call declare_NOTICES_menu_dialog(script_array_NOTICES_main, "HC")
    ElseIf ButtonPressed = SNAP_button then
        Call declare_NOTICES_menu_dialog(script_array_NOTICES_main, "SNAP")
    ElseIf ButtonPressed = CASH_button then
        Call declare_NOTICES_menu_dialog(script_array_NOTICES_main, "CASH")
    ElseIf ButtonPressed = WORD_DOCS_button then
        Call declare_NOTICES_menu_dialog(script_array_NOTICES_main, "WORD")
	End if

	dialog NOTICES_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the Instructions
	IF buttonpressed = Instructions_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/sa/ews/BlueZone_Script_Instructions/Forms/AllItems.aspx?RootFolder=%2Fhsphd%2Fsa%2Fews%2FBlueZone%5FScript%5FInstructions%2FNOTICES&FolderCTID=0x012000A05B86818A1703428050D2E34B3E8EA1&View=%7BFFD55BF9%2D6CDF%2D4B5C%2DB47B%2D3701445A9B34%7D")
Loop until 	ButtonPressed <> Instructions_button and _
			ButtonPressed <> NOTICES_main_button and _
            ButtonPressed <> HEALTH_CARE_button and _
            ButtonPressed <> SNAP_button and _
            ButtonPressed <> CASH_button and _
            ButtonPressed <> WORD_DOCS_button

'MsgBox buttonpressed = script_array_NOTICES_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
If ButtonPressed = ADD_WCOM_button then call run_from_GitHub(script_repository & "notices/add-wcom.vbs")
For i = 0 to ubound(script_array_NOTICES_main)
	If ButtonPressed = script_array_NOTICES_main(i).button then call run_from_GitHub(script_repository & "notices/" & script_array_NOTICES_main(i).file_name)
Next

' For i = 0 to ubound(script_array_NOTICES_list)
' 	If ButtonPressed = script_array_NOTICES_list(i).button then call run_from_GitHub(script_repository & "notices/" & script_array_NOTICES_list(i).file_name)
' Next

stopscript
