'STATS GATHERING=============================================================================================================
name_of_script = "NOTES - QC RESULTS.vbs"       'Replace TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"

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
call changelog_update("06/01/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS FOR THE SCRIPT======================================================================================================
BeginDialog case_number_dialog, 0, 0, 126, 90, "Enter the initial case information"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 100, 25, 20, 15, MAXIS_footer_year
  DropListBox 75, 45, 45, 15, "Select One..."+chr(9)+"CAPER"+chr(9)+"QC"+chr(9)+"QC< $38", error_type
  ButtonGroup ButtonPressed
    OkButton 15, 65, 50, 15
    CancelButton 70, 65, 50, 15
  Text 20, 10, 55, 10, "Case Number: "
  Text 5, 30, 65, 10, "Footer month/year:"
  Text 30, 45, 40, 10, "Error Type:"
EndDialog

BeginDialog Dialog1, 0, 0, 191, 105, "Dialog"
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  DropListBox 20, 60, 150, 15, "Select one..."+chr(9)+"Abdi Ugas"+chr(9)+"Chi Yang"+chr(9)+"Erin Good"+chr(9)+"Gary Lesney"+chr(9)+"Khin Win"+chr(9)+"Lisa Enstad"+chr(9)+"Lor Yang"+chr(9)+"Lori Bona"+chr(9)+"Yer Yang", QC_contact
EndDialog

'Name 	Phone
'Abdi Ugas	651-431-4026
'Chi Yang	651-431-3964
'Erin Good	651-431-3984
'Gary Lesney	651-431-3983
'Khin Win	651-431-5609
'Lisa Enstad	651-431-4115
'Lor Yang	651-431-6304
'Lori Bona	651-431-3950
'Yer Yang	651-431-3965

If QC_contact = "Abdi Ugas"     then phone_number = "651-431-4026"
If QC_contact = "Chi Yang"      then phone_number = "651-431-3964"
If QC_contact = "Erin Good"     then phone_number = "651-431-3984"
If QC_contact = "Gary Lesney"   then phone_number = "651-431-3983"
If QC_contact = "Khin Win"      then phone_number = "651-431-5609"
If QC_contact = "Lisa Enstad"   then phone_number = "651-431-4115"
If QC_contact = "Lor Yang"      then phone_number = "651-431-6304"
If QC_contact = "Lori Bona"     then phone_number = "651-431-3950"
If QC_contact = "Yer Yang"      then phone_number = "651-431-3965"

'END DIALOGS=================================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect ""    'Connects to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabs the MAXIS case number automatically
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'the dialog
Do
	Do
  		err_msg = ""
  		Dialog case_number_dialog
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If error_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select the error type."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in		

Call MAXIS_footer_month_confirmation


MAXIS_background_check

''Checks Maxis for password prompt
'CALL check_for_MAXIS(True)
'
''Now it navigates to a blank case note
'start_a_blank_case_note
'
''...and enters a title (replace variables with your own content)...
'CALL write_variable_in_case_note("*** CASE NOTE HEADER ***")
'
''...some editboxes or droplistboxes (replace variables with your own content)...
'CALL write_bullet_and_variable_in_case_note( "Here's the first bullet",  a_variable_from_your_dialog        )
'CALL write_bullet_and_variable_in_case_note( "Here's another bullet",    another_variable_from_your_dialog  )
'
''...checkbox responses (replace variables with your own content)...
'If some_checkbox_from_your_dialog = checked     then CALL write_variable_in_case_note( "* The checkbox was checked."     )
'If some_checkbox_from_your_dialog = unchecked   then CALL write_variable_in_case_note( "* The checkbox was not checked." )
'
''...and a worker signature. 
'CALL write_variable_in_case_note("---")
'CALL write_variable_in_case_note(worker_signature)
'
'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("")
