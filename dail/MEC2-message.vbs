'STATS GATHERING=============================================================================================================
name_of_script = "DAIL - MEC2 Message.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("01/01/01", "Initial version.", "Ilse Ferris, Hennepin County") 'REPLACE with release date and your name.

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    				'Grabs the MAXIS case number automatically
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)		'Grabs the footer month and year from MAXIS
'  -- OR --
MAXIS_footer_month = CM_plus_1_mo									'Directly assigns a footer month based on the current month
MAXIS_footer_year = CM_plus_1_yr

Dialog1 = "" 'blanking out dialog name
'Add dialog here: Add the dialog just before calling the dialog below unless you need it in the dialog due to using COMBO Boxes or other looping reasons. Blank out the dialog name with Dialog1 = "" before adding dialog.
'    Some Dialog Elements:  Initial Dialog Header: 			BeginDialog Dialog1, 0, 0, 191, 105, "CATEGORY - NAME Case Number Dialog"  				-- Use CATEGORY - NAME somewhere in the header
'							Script Instructions Button:		PushButton 135, 5, 50, 15, "Instructions", script_instructions_btn						-- Have a button to open the instructions
'							Script Purpose/Overview: 		Text 10, 70, 120, 30, "Here is a quick summary of the purpose of the script."			-- Give the worker a little guidance
'							Include edit boxes for necessary details like Case Number, Footer Month, Footer Year, and Worker Signature
'Shows dialog (replace "sample_dialog" with the actual dialog you entered above)----------------------------------
DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation or cancel_without_confirmation
        'Add in all of your mandatory field handling from your dialog here.
        Call validate_MAXIS_case_number(err_msg, "*") ' IF NEEDED
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")   'IF NEEDED
        'The rest of the mandatory handling here
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note." 'IF NEEDED
		If ButtonPressed = script_instructions_btn Then
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/CATEGORY/CATEGORY%20-%20NAME.docx"	'copy the instructions URL here
			err_msg = "LOOP"
		End If
        IF err_msg <> "" AND err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    'Add to all dialogs where you need to work within BLUEZONE
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'End dialog section-----------------------------------------------------------------------------------------------

'Checks to see if in MAXIS
CALL check_for_MAXIS(True) or Call check_for_MAXIS(False)

'Reset to SELF to check the MAXIS region
'This is also helpful to ensure we are not starting in a CASE/NOTE or something
Call back_to_SELF
Call clear_line_of_text(18, 43) 					'clear and rewrite the CASE Number. This is optional but can help the worker not to lose the case number.
EMWriteScreen MAXIS_case_number, 18, 43             'writing in the case number so that if cancelled, the worker doesn't lose the case number.

'MAXIS Region Check
'OPTIONAL - there may be a good reason to be able to run in inquiry or production
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then
		Call script_end_procedure("~PT NAME Script cancelled as it was run in inquiry.")
	End If
End If

'PRIV Handling
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
If is_this_PRIV = True then script_end_procedure("This case is privileged and you do not have access to it. The script will now end.")

'Out of County Handling
'There are a few reasons to allow a script to run on an out of county case - so review if this is needed.
EMReadScreen pw_county_code, 2, 21, 19
If pw_county_code <> "27" Then script_end_procedure("This case is not in Hennepin County and cannot be updated. The script will now end.")

'Do you need to check for PRIV status
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB")

'Do you need to check to see if case is out of county? Add Out-of-County handling here:
'All your other navigation, data catpure and logic here. any other logic or pre case noting actions here.

Call MAXIS_background_check 'IF NEEDED: meaning if you send it through background. Move this to where it makes sense.

'Do you need to set a TIKL?
Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)

'Now it navigates to a blank case note
Call start_a_blank_case_note

'...and enters a title (replace variables with your own content)...
CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")

'...some editboxes or droplistboxes (replace variables with your own content)...
CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )

'...checkbox responses (replace variables with your own content)...
If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )

'...and a worker signature.
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)
'leave the case note open and in edit mode unless you have a business reason not to (BULK scripts, multiple case notes, etc.)

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")

'Add your closing issue documentation here. Make sure it's the most up-to-date version (date on file).
