'Required for statistical purposes===============================================================================
name_of_script = "NAV - VIEW PNLP.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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
call changelog_update("04/17/2019", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""            'connect to MAXIS

CALL MAXIS_case_number_finder(MAXIS_case_number)                    'autofilling MAXIS case number and footer month/year
CALL MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

BeginDialog Dialog1, 0, 0, 126, 80, "Case Number Dialog"            'The dialog
  EditBox 60, 10, 60, 15, MAXIS_case_number
  EditBox 85, 30, 15, 15, MAXIS_footer_month
  EditBox 105, 30, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 55, 40, 15
    CancelButton 80, 55, 40, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 10, 35, 65, 10, "Footer Month/Year:"
EndDialog

'displaying the dialog to confirm or set the case number and footer month/year
Do
    Do
        err_msg = ""

        dialog Dialog1

        cancel_without_confirmation                         'power the cancel button
        CALL validate_MAXIS_case_number(err_msg, "*")       'making sure the case number is present and valid
        MAXIS_footer_month = trim(MAXIS_footer_month)       'validating the footer month and year
        MAXIS_footer_year = trim(MAXIS_footer_year)

        If IsNumeric(MAXIS_footer_month) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
        If IsNumeric(MAXIS_footer_year) = False Then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."

        If err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg      'showing the error message

    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)          'password handling
Loop until are_we_passworded_out = FALSE


CALL back_to_SELF           'getting out of the case to make sure we go to the right place
Do                          'making sure we are in STAT
    CALL navigate_to_MAXIS_screen("STAT", "SUMM")
    EMReadScreen summ_check, 4, 2, 46
Loop until summ_check = "SUMM"

EMWriteScreen "PNLP", 20, 71        'going to PNLP
transmit

Do
    EMGetCursor row, col            'seeing where the cursor is to start (it it is at 20 there are no panels on the particular page)
    Do while row < 20               'If the row is above row 20 then we should write a 'V' for view
        EMSendKey "V"               'Sending 'V' will automatically move the cursor to the next line
        EMGetCursor row, col        'seeing where we are now for the next loop
    Loop
    transmit                        'once we get to line 20, we need to transmit to get to the next page

    EMReadScreen first_panel, 4, 2, 44      'reading the panel name at the top - when we get to ADDR, then we've queued up all the panels to view.
Loop until first_panel = "ADDR"

script_end_procedure("")            'all done
