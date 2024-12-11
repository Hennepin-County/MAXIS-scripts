'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
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
call changelog_update("12/10/2024", "Fixed bug in date-based ABAWD evaluation to work dynamically by footer month/year selected.", "Ilse Ferris, Hennepin County")
call changelog_update("10/16/2024", "Updated exemption finder to include all current Time-Limited and Work Rules exemptions, and updated user experience (no longer need to select HH members, increased readability, etc.).", "Ilse Ferris, Hennepin County")
call changelog_update("08/19/2019", "Updated script so that if started from the ABAWD Tracking Record pop-up on WREG, the script will read where the cursor is placed in the tracking record and if placed on a specific month, the script will autofill that footer month.", "Casey Love, Hennepin County")
call changelog_update("05/07/2018", "Updated universal ABWAWD function.", "Ilse Ferris, Hennepin County")
call changelog_update("04/25/2018", "Updated SCHL exemption coding.", "Ilse Ferris, Hennepin County")
call changelog_update("04/16/2018", "Updated output of potential exemptions for readability.", "Ilse Ferris, Hennepin County")
call changelog_update("04/10/2018", "Enhanced to check cases coded for homelessness for the 'Unfit for Employment' expansion. Also removed code that checked for SSI applying/appealing as this is no longer an exemption reason.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The Script
EMConnect ""
Call check_for_MAXIS(False)

EMReadScreen are_we_at_ABAWD_tracking_record, 8, 1, 71
If are_we_at_ABAWD_tracking_record = "FMCDJAMF" Then
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
BeginDialog Dialog1, 0, 0, 221, 50, "ACTIONS - ABAWD FSET EXEMPTION CHECK"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 170, 5, 20, 15, MAXIS_footer_month
  EditBox 195, 5, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    PushButton 20, 30, 80, 15, "Script Instructions", script_instructions
    OkButton 120, 30, 45, 15
    CancelButton 170, 30, 45, 15
  Text 105, 10, 65, 10, "Footer month/year:"
  Text 5, 10, 45, 10, "Case number:"
EndDialog

Do
	DO
		err_msg = ""
		dialog Dialog1
		Cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
        If ButtonPressed = script_instructions then 
            call open_URL_in_browser("https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20ABAWD%20FSET%20EXEMPTION%20CHECK.docx")
            err_msg = "LOOP" & err_msg
        End if
		IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

ABAWD_eval_date = MAXIS_footer_month & "/1/" & MAXIS_footer_year

Call check_for_MAXIS(False)
'Confirming that the footer month from the dialog matches the footer month in MAXIS
Call MAXIS_footer_month_confirmation
Call ABAWD_FSET_exemption_finder

script_end_procedure_with_error_report("")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 05/23/2024
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/16/2024
'--Tab orders reviewed & confirmed----------------------------------------------10/16/2024
'--Mandatory fields all present & Reviewed--------------------------------------10/16/2024
'--All variables in dialog match mandatory fields-------------------------------10/16/2024
'--Review dialog names for content and content fit in dialog--------------------10/16/2024
'--FIRST DIALOG--NEW EFF 5/23/2024----------------------------------------------
'--Include script category and name somewhere on first dialog-------------------10/16/2024
'--Create a button to reference instructions------------------------------------10/16/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------10/16/2024-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------10/16/2024-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------10/16/2024-------------------N/A
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used ----10/16/2024-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/16/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------10/16/2024-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------10/16/2024-------------------In function
'--Out-of-County handling reviewed----------------------------------------------10/16/2024-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/16/2024
'--BULK - review output of statistics and run time/count (if applicable)--------10/16/2024-------------------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/16/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/16/2024
'--Incrementors reviewed (if necessary)-----------------------------------------10/16/2024-----------------In function
'--Denomination reviewed -------------------------------------------------------10/16/2024
'--Script name reviewed---------------------------------------------------------10/16/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/16/2024-----------------In function

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/16/2024
'--comment Code-----------------------------------------------------------------10/16/2024
'--Update Changelog for release/update------------------------------------------10/16/2024
'--Remove testing message boxes-------------------------------------------------10/16/2024
'--Remove testing code/unnecessary code-----------------------------------------10/16/2024
'--Review/update SharePoint instructions----------------------------------------10/16/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------10/16/2024
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/16/2024
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------10/16/2024
'--Complete misc. documentation (if applicable)---------------------------------10/16/2024
'--Update project team/issue contact (if applicable)----------------------------10/16/2024
