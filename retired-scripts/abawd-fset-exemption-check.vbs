'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 98                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

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
call changelog_update("08/19/2019", "Updated script so that if started from the ABAWD Tracking Record pop-up on WREG, the script will read where the cursor is placed in the tracking record and if placed on a specific month, the script will autofill that footer month.", "Casey Love, Hennepin County")
call changelog_update("05/07/2018", "Updated universal ABWAWD function.", "Ilse Ferris, Hennepin County")
call changelog_update("04/25/2018", "Updated SCHL exemption coding.", "Ilse Ferris, Hennepin County")
call changelog_update("04/16/2018", "Updated output of potential exemptions for readability.", "Ilse Ferris, Hennepin County")
call changelog_update("04/10/2018", "Enhanced to check cases coded for homelessness for the 'Unfit for Employment' expansion. Also removed code that checked for SSI applying/appealing as this is no longer an exemption reason.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script----------------------------------------------------------------------------------------------------
'Connecting to MAXIS, and grabbing the case number and current footer month/year
EMConnect ""

EMReadScreen are_we_at_ABAWD_tracking_record, 21, 4, 34
If are_we_at_ABAWD_tracking_record = "ABAWD Tracking Record" Then
    EMGetCursor tracker_row, tracker_col

    If tracker_col = 19 Then
        MAXIS_footer_month = "01"
    ElseIf tracker_col = 23 Then
        MAXIS_footer_month = "02"
    ElseIf tracker_col = 27 Then
        MAXIS_footer_month = "03"
    ElseIf tracker_col = 31 Then
        MAXIS_footer_month = "04"
    ElseIf tracker_col = 35 Then
        MAXIS_footer_month = "05"
    ElseIf tracker_col = 39 Then
        MAXIS_footer_month = "06"
    ElseIf tracker_col = 43 Then
        MAXIS_footer_month = "07"
    ElseIf tracker_col = 47 Then
        MAXIS_footer_month = "08"
    ElseIf tracker_col = 51 Then
        MAXIS_footer_month = "09"
    ElseIf tracker_col = 55 Then
        MAXIS_footer_month = "10"
    ElseIf tracker_col = 59 Then
        MAXIS_footer_month = "11"
    ElseIf tracker_col = 63 Then
        MAXIS_footer_month = "12"
    End If

    If MAXIS_footer_month <> "" Then EMReadScreen MAXIS_footer_year, 2, tracker_row, 15

    MX_mo = MAXIS_footer_month * 1
    MX_yr = MAXIS_footer_year * 1
    curr_mo = CM_plus_1_mo * 1
    curr_yr = CM_plus_1_yr * 1

    If  MX_yr > curr_yr Then
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
    ElseIf MX_yr = curr_yr AND MX_mo > curr_mo Then
        MAXIS_footer_month = ""
        MAXIS_footer_year = ""
    End If

    PF3
End If

CALL MAXIS_case_number_finder(MAXIS_case_number)
If MAXIS_footer_month = "" Then call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 166, 70, "Case number dialog"
  EditBox 65, 5, 70, 15, MAXIS_case_number
  EditBox 65, 25, 30, 15, MAXIS_footer_month
  EditBox 130, 25, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 35, 50, 50, 15
    CancelButton 95, 50, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 100, 30, 25, 10, "Year:"
EndDialog
Do
	DO
		err_msg = ""
		dialog Dialog1
		cancel_confirmation
		IF MAXIS_case_number = "" THEN err_msg = err_msg & vbCr & "* Please enter a case number."
		IF MAXIS_footer_month = "" THEN err_msg = err_msg & vbCr & "* Please enter a benefit month."
		IF MAXIS_footer_year = "" THEN err_msg = err_msg & vbCr & "* Please enter a benefit year."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Confirming that the footer month from the dialog matches the footer month in MAXIS
Call MAXIS_footer_month_confirmation
Call ABAWD_FSET_exemption_finder

script_end_procedure("")
